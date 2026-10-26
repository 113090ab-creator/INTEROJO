import json
import sys
import unittest
from pathlib import Path
from unittest import mock

import requests


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import app  # noqa: E402


def response(status: int, retry_after: str = "") -> requests.Response:
    result = requests.Response()
    result.status_code = status
    result._content = json.dumps({"rows": [1]}).encode("utf-8")
    if retry_after:
        result.headers["Retry-After"] = retry_after
    return result


class PlanApiRetryTests(unittest.TestCase):
    def request(self, responses: list[requests.Response]):
        with (
            mock.patch.object(app, "PLAN_API_RETRY_ATTEMPTS", 3),
            mock.patch.object(app.requests, "get", side_effect=responses) as get,
            mock.patch.object(app.time, "sleep") as sleep,
        ):
            result = app.request_plan_api_payload("https://api.example.test/data", {}, "test-key")
        return result, get, sleep

    def test_rate_limit_waits_before_retrying_and_recovers(self) -> None:
        result, get, sleep = self.request([response(429), response(429), response(200)])
        self.assertEqual(result, ({"rows": [1]}, ""))
        self.assertEqual(get.call_count, 3)
        self.assertEqual(sleep.call_args_list, [mock.call(10.0), mock.call(20.0)])

    def test_retry_after_seconds_controls_retry(self) -> None:
        result, _, sleep = self.request([response(429, "30"), response(200)])
        self.assertEqual(result, ({"rows": [1]}, ""))
        sleep.assert_called_once_with(30.0)

    def test_retry_after_http_date_controls_retry(self) -> None:
        with mock.patch.object(app.time, "time", return_value=1792980000.0):
            delay = app.get_plan_api_retry_delay(response(429, "Mon, 26 Oct 2026 02:00:30 GMT"), 0)
        self.assertEqual(delay, 30.0)

    def test_invalid_retry_after_uses_rate_limit_backoff(self) -> None:
        for value in ("invalid", "-1", "nan"):
            with self.subTest(retry_after=value):
                self.assertEqual(app.get_plan_api_retry_delay(response(429, value), 1), 20.0)

    def test_other_transient_errors_keep_short_retry(self) -> None:
        result, _, sleep = self.request([response(503), response(200)])
        self.assertEqual(result, ({"rows": [1]}, ""))
        sleep.assert_called_once_with(0.4)

    def test_permanent_errors_are_not_retried(self) -> None:
        result, get, sleep = self.request([response(401)])
        self.assertIsNone(result[0])
        self.assertIn("HTTP 401", result[1])
        self.assertEqual(get.call_count, 1)
        sleep.assert_not_called()

    def test_exhausted_rate_limit_returns_error_for_workflow_retry(self) -> None:
        result, get, sleep = self.request([response(429), response(429), response(429)])
        self.assertIsNone(result[0])
        self.assertIn("HTTP 429", result[1])
        self.assertEqual(get.call_count, 3)
        self.assertEqual(sleep.call_count, 2)


if __name__ == "__main__":
    unittest.main()
