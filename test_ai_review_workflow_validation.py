from __future__ import annotations

import unittest
from unittest.mock import patch

from term_extractor_app.ai_review import workflow_service
from term_extractor_app.ai_review.review_service import (
    ReviewTaskError,
    _build_packages,
    _build_user_prompt,
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

    def test_no_source_review_prompt_does_not_invent_an_original_text(self) -> None:
        prompt = _build_user_prompt(
            {"source_language": "none", "target_language": "中文", "user_prompt": "审校：{text}"},
            "{}",
            False,
        )
        self.assertIn("没有原文", prompt)
        self.assertNotIn("none 原文", prompt)

    def test_term_review_placeholder_is_only_expanded_for_matched_terms(self) -> None:
        config = {
            "source_language": "日语",
            "target_language": "简体中文",
            "user_prompt": "规则：\n{term_review}\n数据：{text}",
        }
        without_terms = _build_user_prompt(config, "{}", False, False)
        with_terms = _build_user_prompt(config, "{}", False, True)
        self.assertNotIn("{term_review}", without_terms)
        self.assertNotIn("术语表审校规则", without_terms)
        self.assertIn("术语表审校规则", with_terms)

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

    def test_conversation_preview_only_returns_latest_workspace_run(self) -> None:
        snapshot = {
            "workspace_runs": [{"id": "run-new"}, {"id": "run-old"}],
            "tasks": [
                {"task_id": "task-old", "target_language": "ko"},
                {"task_id": "task-new", "target_language": "en"},
            ],
        }
        tasks = {
            "task-old": {"id": "task-old", "batch_id": "batch-old", "output_path": "old.xlsx"},
            "task-new": {"id": "task-new", "batch_id": "batch-new", "output_path": "new.xlsx"},
        }
        batches = {
            "batch-old": {"metadata": {"workspace_run_id": "run-old"}},
            "batch-new": {"metadata": {"workspace_run_id": "run-new"}},
        }
        with (
            patch.object(workflow_service, "get_session_snapshot", return_value=snapshot),
            patch.object(workflow_service, "get_review_task", side_effect=lambda task_id: tasks[task_id]),
            patch.object(workflow_service, "get_batch", side_effect=lambda batch_id: batches[batch_id]),
            patch.object(workflow_service, "get_review_results", return_value=[]),
        ):
            results = workflow_service.get_session_task_results("session")
        self.assertEqual([item["target_language"] for item in results], ["en"])
        self.assertEqual(results[0]["task"]["output_path"], "new.xlsx")

    def test_new_workspace_run_hides_previous_results_before_tasks_exist(self) -> None:
        snapshot = {
            "workspace_runs": [{"id": "run-new"}, {"id": "run-old"}],
            "tasks": [{"task_id": "task-old", "target_language": "ko"}],
        }
        with (
            patch.object(workflow_service, "get_session_snapshot", return_value=snapshot),
            patch.object(workflow_service, "get_review_task", return_value={"id": "task-old", "batch_id": "batch-old"}),
            patch.object(workflow_service, "get_batch", return_value={"metadata": {"workspace_run_id": "run-old"}}),
        ):
            self.assertEqual(workflow_service.get_session_task_results("session"), [])


if __name__ == "__main__":
    unittest.main()
