from __future__ import annotations

import unittest

from term_extractor_app.ai_review.review_service import (
    ReviewTaskError,
    _build_packages,
    _validate_response_items,
)
from term_extractor_app.ai_review.output_service import NORMAL_HEADERS


def _request_item(item_id: str) -> dict[str, str]:
    return {"id": item_id, "source_text": "source", "target_text": "target"}


def _response_item(item_id: str) -> dict[str, object]:
    return {
        "id": item_id,
        "has_issue": False,
        "issue_type": "",
        "issue": "",
        "suggestion": "",
    }


class ReviewWorkflowValidationTests(unittest.TestCase):
    def test_output_includes_filename_column(self) -> None:
        self.assertEqual(NORMAL_HEADERS[0], "文件名")

    def test_packages_respect_item_and_character_limits(self) -> None:
        items = [_request_item(str(index)) for index in range(161)]
        packages = _build_packages(items, 1_000_000, max_items=80)
        self.assertEqual([len(package) for package in packages], [80, 80, 1])

        character_packages = _build_packages(items[:3], 15, max_items=80)
        self.assertEqual([len(package) for package in character_packages], [1, 1, 1])

    def test_validation_rejects_missing_item(self) -> None:
        package = [_request_item("a"), _request_item("b")]
        with self.assertRaises(ReviewTaskError):
            _validate_response_items({"items": [_response_item("a")]}, package, {"mode": "normal"})

    def test_validation_rejects_duplicate_and_unknown_ids(self) -> None:
        package = [_request_item("a"), _request_item("b")]
        with self.assertRaises(ReviewTaskError):
            _validate_response_items(
                {"items": [_response_item("a"), _response_item("a")]},
                package,
                {"mode": "normal"},
            )
        with self.assertRaises(ReviewTaskError):
            _validate_response_items(
                {"items": [_response_item("a"), _response_item("c")]},
                package,
                {"mode": "normal"},
            )

    def test_validation_rejects_missing_required_field(self) -> None:
        item = _response_item("a")
        del item["suggestion"]
        with self.assertRaises(ReviewTaskError):
            _validate_response_items({"items": [item]}, [_request_item("a")], {"mode": "normal"})


if __name__ == "__main__":
    unittest.main()
