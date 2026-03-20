from __future__ import annotations

import base64
import json
import os
import time
from io import BytesIO
from typing import Annotated, List, Optional

import httpx
import PIL
import openai
from openai import APITimeoutError, RateLimitError
from pydantic import BaseModel

from marker.logger import get_logger
from marker.schema.blocks import Block
from marker.services import BaseService

logger = get_logger()


class OpenAICompatJSONService(BaseService):
    """OpenAI-compatible multimodal service for Marker.

    Marker upstream's OpenAIService uses `client.beta.chat.completions.parse(...)` with
    a Pydantic schema. Many OpenAI-compatible proxies (including Gemini gateways)
    don't implement that endpoint/behavior.

    This service instead:
    - calls `client.chat.completions.create(...)`
    - requests JSON-only output via `response_format={"type": "json_object"}`
    - validates against the provided Pydantic schema

    It is intended for OpenAI-compatible servers at `openai_base_url`.
    """

    openai_base_url: Annotated[
        str, "The base url to use for OpenAI-like models. No trailing slash."
    ] = "https://api.openai.com/v1"
    openai_model: Annotated[str, "The model name to use for OpenAI-like model."] = (
        "gpt-5.2"
    )
    openai_api_key: Annotated[str, "The API key to use for the OpenAI-like service."] = (
        None
    )
    openai_image_format: Annotated[
        str, "The image format to use for the OpenAI-like service. Use 'png' for better compatibility."
    ] = "png"
    _http_client: Optional[httpx.Client] = None
    _client: Optional[openai.OpenAI] = None
    _prefer_responses_api: bool = False

    def _image_part(self, img: PIL.Image.Image) -> dict:
        img_fmt = self.openai_image_format
        image_bytes = BytesIO()
        img.save(image_bytes, format=img_fmt)
        b64 = base64.b64encode(image_bytes.getvalue()).decode("utf-8")
        return {
            "type": "image_url",
            "image_url": {"url": f"data:image/{img_fmt.lower()};base64,{b64}"},
        }

    def process_images(self, images: List[PIL.Image.Image]) -> List[dict]:
        if isinstance(images, PIL.Image.Image):
            images = [images]
        return [self._image_part(img) for img in images]

    def _pick_http_proxy(self) -> Optional[str]:
        for key in ("HTTPS_PROXY", "https_proxy", "HTTP_PROXY", "http_proxy"):
            value = os.environ.get(key)
            if not value:
                continue
            proxy = value.strip()
            if proxy.lower().startswith(("http://", "https://")):
                return proxy
        return None

    def _has_socks_all_proxy(self) -> bool:
        for key in ("ALL_PROXY", "all_proxy"):
            value = os.environ.get(key)
            if value and value.strip().lower().startswith("socks"):
                return True
        return False

    def get_client(self) -> openai.OpenAI:
        if self._client is not None:
            return self._client

        # Avoid inheriting ALL_PROXY=socks5://... when socksio is not installed.
        # If an HTTP(S) proxy is configured, use it explicitly.
        proxy = self._pick_http_proxy()
        if self._has_socks_all_proxy() and proxy is None:
            logger.warning(
                "Detected SOCKS ALL_PROXY but no HTTP(S) proxy override; ignoring env proxy for Marker LLM requests"
            )

        self._http_client = httpx.Client(trust_env=False, proxy=proxy)
        self._client = openai.OpenAI(
            api_key=self.openai_api_key,
            base_url=self.openai_base_url,
            http_client=self._http_client,
        )
        return self._client

    def _should_use_responses_api(self, message: str) -> bool:
        msg = message.lower()
        return (
            "unsupported parameter: messages" in msg
            or "stream must be set to true" in msg
            or "unsupported legacy protocol" in msg
            or ("/v1/chat/completions" in msg and "/v1/responses" in msg)
        )

    def _to_responses_content(self, image_data: List[dict], strict_prompt: str) -> list[dict]:
        content: list[dict] = []
        for part in image_data:
            if part.get("type") != "image_url":
                continue
            image_url = part.get("image_url")
            if isinstance(image_url, dict):
                url = image_url.get("url")
                if isinstance(url, str) and url:
                    content.append({"type": "input_image", "image_url": url})
        content.append({"type": "input_text", "text": strict_prompt})
        return content

    def _create_with_responses_api(
        self,
        *,
        client: openai.OpenAI,
        image_data: List[dict],
        strict_prompt: str,
        timeout: int,
    ) -> tuple[str, Optional[int]]:
        content = self._to_responses_content(image_data, strict_prompt)
        stream = client.responses.create(
            extra_headers={
                "X-Title": "Marker",
                "HTTP-Referer": "https://github.com/datalab-to/marker",
            },
            model=self.openai_model,
            input=[{"role": "user", "content": content}],
            timeout=timeout,
            stream=True,
            max_output_tokens=self.max_output_tokens,
        )

        chunks: list[str] = []
        total_tokens: Optional[int] = None
        for event in stream:
            event_type = getattr(event, "type", "")
            if event_type == "response.output_text.delta":
                delta = getattr(event, "delta", "")
                if isinstance(delta, str):
                    chunks.append(delta)
            elif event_type == "response.completed":
                response = getattr(event, "response", None)
                usage = getattr(response, "usage", None)
                total_tokens = getattr(usage, "total_tokens", None)

        return "".join(chunks) or "{}", total_tokens

    def _fix_invalid_json_escapes(self, s: str) -> str:
        # Fix common invalid backslash escapes inside JSON strings.
        # Example: "C:\path\file" is valid, but "C:\path\new\t" might contain "\p" etc.
        valid_escapes = {"\"", "\\", "/", "b", "f", "n", "r", "t"}

        out: list[str] = []
        in_str = False
        i = 0
        while i < len(s):
            ch = s[i]
            if ch == '"':
                # toggle string context if not escaped
                backslashes = 0
                j = i - 1
                while j >= 0 and s[j] == "\\":
                    backslashes += 1
                    j -= 1
                if backslashes % 2 == 0:
                    in_str = not in_str
                out.append(ch)
                i += 1
                continue

            if in_str and ch == "\\" and i + 1 < len(s):
                nxt = s[i + 1]
                if nxt in valid_escapes:
                    out.append(ch)
                    out.append(nxt)
                    i += 2
                    continue
                if nxt == "u" and i + 5 < len(s):
                    hex_part = s[i + 2 : i + 6]
                    if all(c in "0123456789abcdefABCDEF" for c in hex_part):
                        out.append(ch)
                        out.append(nxt)
                        out.append(hex_part)
                        i += 6
                        continue
                # invalid escape: turn `\X` into `\\X`
                out.append("\\\\")
                i += 1
                continue

            out.append(ch)
            i += 1

        return "".join(out)

    def _json_loads_robust(self, s: str):
        s2 = s.strip()
        decoder = json.JSONDecoder()
        try:
            return json.loads(s2)
        except json.JSONDecodeError as e:
            # Some providers may append extra text after JSON.
            try:
                obj, _idx = decoder.raw_decode(s2.lstrip())
                return obj
            except json.JSONDecodeError:
                pass

            if "Invalid \\escape" in str(e):
                fixed = self._fix_invalid_json_escapes(s2)
                try:
                    return json.loads(fixed)
                except json.JSONDecodeError:
                    obj, _idx = decoder.raw_decode(fixed.lstrip())
                    return obj

            raise

    def _extract_json_data(self, text: str):
        text_stripped = text.strip()
        try:
            return self._json_loads_robust(text_stripped)
        except Exception:
            start = text_stripped.find("{")
            end = text_stripped.rfind("}")
            if start == -1 or end == -1 or end <= start:
                raise
            return self._json_loads_robust(text_stripped[start : end + 1])

    def __call__(
        self,
        prompt: str,
        image: PIL.Image.Image | List[PIL.Image.Image] | None,
        block: Block | None,
        response_schema: type[BaseModel],
        max_retries: int | None = None,
        timeout: int | None = None,
    ):
        if max_retries is None:
            max_retries = self.max_retries
        if timeout is None:
            timeout = self.timeout

        client = self.get_client()
        image_data = self.format_image_for_llm(image)

        schema_json = response_schema.model_json_schema()
        strict_prompt = (
            "You MUST output a single JSON object only (no Markdown, no code fences).\n"
            "The JSON must match this JSON schema exactly.\n\n"
            f"JSON schema: {json.dumps(schema_json, ensure_ascii=True)}\n\n"
            "Task:\n"
            f"{prompt}"
        )

        messages = [
            {
                "role": "user",
                "content": [
                    *image_data,
                    {"type": "text", "text": strict_prompt},
                ],
            }
        ]

        total_tries = max_retries + 1
        for tries in range(1, total_tries + 1):
            try:
                if self._prefer_responses_api:
                    text, total_tokens = self._create_with_responses_api(
                        client=client,
                        image_data=image_data,
                        strict_prompt=strict_prompt,
                        timeout=timeout,
                    )
                    data = self._extract_json_data(text)
                    validated = response_schema.model_validate(data)
                    if block and total_tokens is not None:
                        block.update_metadata(llm_tokens_used=total_tokens, llm_request_count=1)
                    return validated.model_dump()

                def _create_chat_completion(*, use_response_format: bool):
                    kwargs = {
                        "extra_headers": {
                            "X-Title": "Marker",
                            "HTTP-Referer": "https://github.com/datalab-to/marker",
                        },
                        "model": self.openai_model,
                        "messages": messages,
                        "timeout": timeout,
                    }
                    if use_response_format:
                        # Some OpenAI-compatible providers/models don't support `response_format`.
                        kwargs["response_format"] = {"type": "json_object"}
                    return client.chat.completions.create(**kwargs)

                total_tokens: Optional[int] = None
                try:
                    resp = _create_chat_completion(use_response_format=True)
                except Exception as e:
                    # Volcengine `ark-code-latest` (and some other OpenAI-compatible models)
                    # rejects `response_format.type=json_object`.
                    msg = str(e)
                    if (
                        "response_format" in msg
                        and "json_object" in msg
                        and ("not supported" in msg or "InvalidParameter" in msg)
                    ):
                        logger.warning(
                            "Provider/model does not support response_format=json_object; retrying without it"
                        )
                        try:
                            resp = _create_chat_completion(use_response_format=False)
                        except Exception as e2:
                            msg2 = str(e2)
                            if self._should_use_responses_api(msg2):
                                self._prefer_responses_api = True
                                logger.debug(
                                    "Provider requires responses API; switching Marker LLM calls to responses mode"
                                )
                                text, total_tokens = self._create_with_responses_api(
                                    client=client,
                                    image_data=image_data,
                                    strict_prompt=strict_prompt,
                                    timeout=timeout,
                                )
                            else:
                                raise
                    else:
                        if self._should_use_responses_api(msg):
                            self._prefer_responses_api = True
                            logger.debug(
                                "Provider requires responses API; switching Marker LLM calls to responses mode"
                            )
                            text, total_tokens = self._create_with_responses_api(
                                client=client,
                                image_data=image_data,
                                strict_prompt=strict_prompt,
                                timeout=timeout,
                            )
                        else:
                            raise
                else:
                    text = resp.choices[0].message.content or "{}"
                    usage = getattr(resp, "usage", None)
                    total_tokens = getattr(usage, "total_tokens", None)

                data = self._extract_json_data(text)

                validated = response_schema.model_validate(data)

                if block and total_tokens is not None:
                    block.update_metadata(llm_tokens_used=total_tokens, llm_request_count=1)

                return validated.model_dump()

            except (APITimeoutError, RateLimitError) as e:
                if tries == total_tries:
                    logger.error(
                        f"Rate limit/timeout error: {e}. Max retries reached. Giving up. (Attempt {tries}/{total_tries})"
                    )
                    break
                wait_time = tries * self.retry_wait_time
                logger.warning(
                    f"Rate limit/timeout error: {e}. Retrying in {wait_time} seconds... (Attempt {tries}/{total_tries})"
                )
                time.sleep(wait_time)
            except Exception as e:
                logger.error(f"OpenAI-compatible inference failed: {e}")
                break

        return {}
