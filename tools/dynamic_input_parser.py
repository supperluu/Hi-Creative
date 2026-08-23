import json
import re
from urllib.parse import urlparse


def main(input_value) -> dict:
    """Parse batch data from plain text, a JSON string, dict, or list."""

    def fail(message, errors=None):
        return {
            "batch_ok": False,
            "parse_error": message,
            "row_count": 0,
            "valid_record_count": 0,
            "total_image_count": 0,
            "downloadable_image_count": 0,
            "image_urls": [],
            "batch_input_json": "{}",
            "row_errors_json": json.dumps(errors or [], ensure_ascii=False),
            "empty_results": [],
        }

    def try_json(value):
        if not isinstance(value, str):
            return None
        candidate = value.strip().lstrip("\ufeff")
        if not candidate or candidate[0] not in "[{":
            return None
        try:
            parsed = json.loads(candidate)
        except (TypeError, ValueError):
            return None
        return parsed if isinstance(parsed, (dict, list)) else None

    def is_payload(value):
        """Recognize a record by its structure, not by its wrapper key."""
        return (
            isinstance(value, dict)
            and isinstance(value.get("question"), dict)
            and isinstance(value.get("answer"), dict)
        )

    def find_payloads(value, found, seen):
        """Recursively unwrap arbitrary objects and JSON strings."""
        if isinstance(value, str):
            parsed = try_json(value)
            if parsed is not None:
                find_payloads(parsed, found, seen)
            return

        if isinstance(value, dict):
            if is_payload(value):
                object_id = id(value)
                if object_id not in seen:
                    seen.add(object_id)
                    found.append(value)
                return
            for child in value.values():
                find_payloads(child, found, seen)
            return

        if isinstance(value, list):
            for child in value:
                find_payloads(child, found, seen)

    def find_text_candidates(value, found, depth=0):
        """Find document-extraction text without assuming a wrapper field name."""
        if depth > 20:
            return
        if isinstance(value, str):
            parsed = try_json(value)
            if parsed is not None:
                find_text_candidates(parsed, found, depth + 1)
            else:
                candidate = value.strip().lstrip("\ufeff")
                if candidate:
                    found.append(candidate)
            return
        if isinstance(value, dict):
            for child in value.values():
                find_text_candidates(child, found, depth + 1)
            return
        if isinstance(value, list):
            for child in value:
                find_text_candidates(child, found, depth + 1)

    def rows_from_text(text):
        rows = []

        # Dify document extractor: Markdown pipe table.
        markdown_header = re.search(
            r"(?m)^\s*\|\s*数据ID\s*\|\s*payload_json\s*\|\s*$", text
        )
        if markdown_header:
            markdown_row = re.compile(
                r"^\s*\|\s*([^|]+?)\s*\|\s*(\{.*\})\s*\|\s*$"
            )
            for line_number, line in enumerate(text.splitlines(), 1):
                match = markdown_row.match(line)
                if not match:
                    continue
                csv_data_id = match.group(1).strip()
                if csv_data_id in {"数据ID", "----"}:
                    continue
                rows.append(
                    {
                        "row_number": line_number,
                        "csv_data_id": csv_data_id,
                        "payload_raw": match.group(2).strip(),
                    }
                )

        # Another Dify extraction form.
        if not rows:
            marker = re.compile(
                r"(?m)^\s*数据ID\s*:\s*(.*?)\s*;\s*payload_json\s*:\s*"
            )
            matches = list(marker.finditer(text))
            for index, match in enumerate(matches):
                end = (
                    matches[index + 1].start()
                    if index + 1 < len(matches)
                    else len(text)
                )
                rows.append(
                    {
                        "row_number": index + 2,
                        "csv_data_id": match.group(1).strip(),
                        "payload_raw": text[match.end() : end].strip(),
                    }
                )
        return rows

    # First try structured input. It can be a dict/list or a JSON string, and can
    # have any number of wrapper levels with any field names.
    structured_input = try_json(input_value) if isinstance(input_value, str) else input_value
    payloads = []
    if isinstance(structured_input, (dict, list)):
        find_payloads(structured_input, payloads, set())

    if payloads:
        raw_rows = [
            {
                "row_number": index,
                "csv_data_id": str(payload.get("dataId") or ""),
                "payload": payload,
            }
            for index, payload in enumerate(payloads, 1)
        ]
    else:
        # Otherwise discover the document text by content, not a fixed wrapper key.
        text_candidates = []
        find_text_candidates(input_value, text_candidates)
        raw_rows = []
        for candidate in text_candidates:
            candidate_rows = rows_from_text(candidate)
            if candidate_rows:
                raw_rows = candidate_rows
                break

    if not raw_rows:
        return fail(
            "未识别到数据。输入可以是payload对象、payload对象数组、它们的JSON字符串，"
            "或包含‘数据ID’和‘payload_json’两列的Dify文档提取文本"
        )

    errors = []
    records = []
    image_urls = []
    visual_file_map = []
    total_image_count = 0

    for row in raw_rows:
        row_number = row["row_number"]
        csv_data_id = row["csv_data_id"]

        if "payload" in row:
            payload = row["payload"]
        else:
            try:
                payload = json.loads(row["payload_raw"].strip())
            except Exception as exc:
                errors.append(
                    {
                        "row_number": row_number,
                        "data_id": csv_data_id,
                        "error": "payload_json解析失败: " + str(exc),
                    }
                )
                continue

        if not isinstance(payload, dict):
            errors.append(
                {
                    "row_number": row_number,
                    "data_id": csv_data_id,
                    "error": "payload_json顶层必须是对象",
                }
            )
            continue

        question = payload.get("question") or {}
        answer = payload.get("answer") or {}
        context = payload.get("evaluationContext") or {}
        q_content = question.get("content") or []
        a_content = answer.get("content") or []
        if not isinstance(q_content, list) or not isinstance(a_content, list):
            errors.append(
                {
                    "row_number": row_number,
                    "data_id": csv_data_id,
                    "error": "question.content和answer.content必须是数组",
                }
            )
            continue

        question_texts = []
        for item in q_content:
            if isinstance(item, dict) and item.get("type") == "text":
                question_texts.append(str(item.get("text") or ""))

        normalized_segments = []
        image_items = []
        preceding_text_segment_id = ""
        preceding_text = ""
        answer_texts = []
        warnings = []

        for segment_index, item in enumerate(a_content):
            if not isinstance(item, dict):
                warnings.append(
                    "answer.content第%d项不是对象" % (segment_index + 1)
                )
                continue
            item_type = str(item.get("type") or "")
            segment_id = str(
                item.get("segmentId") or "segment_%d" % (segment_index + 1)
            )
            if item_type == "text":
                segment_text = str(item.get("text") or "")
                answer_texts.append(segment_text)
                preceding_text_segment_id = segment_id
                preceding_text = segment_text
                normalized_segments.append(
                    {
                        "type": "text",
                        "segment_id": segment_id,
                        "text": segment_text,
                        "content_index": segment_index,
                    }
                )
            elif item_type == "image":
                url = str(item.get("downloadUrl") or item.get("url") or "").strip()
                parsed = urlparse(url) if url else None
                url_valid = bool(
                    parsed and parsed.scheme in ("http", "https") and parsed.netloc
                )
                image_entry = {
                    "type": "image",
                    "segment_id": segment_id,
                    "asset_id": str(item.get("assetId") or ""),
                    "position": item.get("position"),
                    "url": url,
                    "url_valid": url_valid,
                    "mime_type": str(item.get("mimeType") or ""),
                    "content_index": segment_index,
                    "paired_text_segment_id": preceding_text_segment_id,
                    "paired_text": preceding_text,
                }
                total_image_count += 1
                if url_valid:
                    visual_index = len(image_urls) + 1
                    image_urls.append(url)
                    image_entry["global_visual_index"] = visual_index
                    visual_file_map.append(
                        {
                            "global_visual_index": visual_index,
                            "row_number": row_number,
                            "data_id": str(payload.get("dataId") or csv_data_id),
                            "case_id": str(payload.get("caseId") or ""),
                            "segment_id": segment_id,
                            "asset_id": image_entry["asset_id"],
                            "position": image_entry["position"],
                            "paired_text_segment_id": preceding_text_segment_id,
                        }
                    )
                else:
                    image_entry["global_visual_index"] = None
                    warnings.append("图片%s缺少有效HTTP(S)地址" % segment_id)
                image_items.append(image_entry)
                normalized_segments.append(image_entry)
            else:
                warnings.append("忽略未知segment类型: %s" % item_type)

        payload_data_id = str(payload.get("dataId") or csv_data_id)
        if csv_data_id and payload_data_id and csv_data_id != payload_data_id:
            warnings.append("CSV数据ID与payload.dataId不一致")
        declared_images = context.get("imageCount")
        if isinstance(declared_images, int) and declared_images != len(image_items):
            warnings.append("evaluationContext.imageCount与实际图片段数量不一致")
        declared_segments = context.get("segmentCount")
        if isinstance(declared_segments, int) and declared_segments != len(a_content):
            warnings.append("evaluationContext.segmentCount与answer.content数量不一致")

        records.append(
            {
                "batch_record_index": len(records) + 1,
                "csv_row_number": row_number,
                "csv_data_id": csv_data_id,
                "schema_version": str(payload.get("schemaVersion") or ""),
                "data_id": payload_data_id,
                "case_id": str(payload.get("caseId") or ""),
                "question_text": "\n".join(question_texts).strip(),
                "answer_text": "\n".join(answer_texts).strip(),
                "answer_segments_in_order": normalized_segments,
                "image_items_in_content_order": image_items,
                "evaluation_context": context,
                "response_date": str(context.get("responseDate") or ""),
                "actual_image_count": len(image_items),
                "downloadable_image_count": sum(
                    1 for item in image_items if item.get("url_valid")
                ),
                "validation_warnings": warnings,
            }
        )

    if errors:
        return fail(
            "输入中存在无法解析的记录；为避免图片与记录错位，已终止整批评测",
            errors,
        )
    if not records:
        return fail("输入中没有可评测的数据记录")

    batch_input = {
        "input_schema": "machine-label-input-v1-dynamic-batch",
        "record_count": len(records),
        "visual_file_count": len(image_urls),
        "visual_file_map_in_exact_order": visual_file_map,
        "records_in_input_order": records,
    }
    return {
        "batch_ok": True,
        "parse_error": "",
        "row_count": len(raw_rows),
        "valid_record_count": len(records),
        "total_image_count": total_image_count,
        "downloadable_image_count": len(image_urls),
        "image_urls": image_urls,
        "batch_input_json": json.dumps(
            batch_input, ensure_ascii=False, separators=(",", ":")
        ),
        "row_errors_json": "[]",
        "empty_results": [],
    }
