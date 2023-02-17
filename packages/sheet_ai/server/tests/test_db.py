from sheet_ai.db import get_prompt_response, get_session_prompt_count, save_prompt_response


def test_db() -> None:
    """Basic sheet_ai.db tests exercising prompts caching and retrieving."""
    assert get_session_prompt_count("session1") == 0
    assert get_session_prompt_count("session2") == 0

    assert get_prompt_response(["prompt1"]) is None
    assert get_prompt_response(["prompt1", "addition1"]) is None
    assert get_prompt_response(["prompt2"]) is None

    save_prompt_response("session1", ["prompt1"], [[{"input": "workbook1"}]])
    save_prompt_response("session1", ["prompt2"], [[{"input": "workbook2"}]])

    assert get_session_prompt_count("session1") == 2
    assert get_session_prompt_count("session2") == 0

    assert get_prompt_response(["prompt1"]) == [[{"input": "workbook1"}]]
    assert get_prompt_response(["prompt1", "addition1"]) is None
    assert get_prompt_response(["prompt2"]) == [[{"input": "workbook2"}]]

    # saving the same prompt doesn't increase the session count
    save_prompt_response("session1", ["prompt1"], [[{"input": "workbook1"}]])
    assert get_session_prompt_count("session1") == 2

    save_prompt_response("session2", ["prompt1", "addition1"], [[{"input": "workbook3"}]])

    assert get_session_prompt_count("session1") == 2
    assert get_session_prompt_count("session2") == 1

    assert get_prompt_response(["prompt1"]) == [[{"input": "workbook1"}]]
    assert get_prompt_response(["prompt1", "addition1"]) == [[{"input": "workbook3"}]]
    assert get_prompt_response(["prompt2"]) == [[{"input": "workbook2"}]]
