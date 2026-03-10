def test_intentional_failure():
    """Intentionally failing test to verify CI healing behavior."""
    assert False, "This test is expected to fail to trigger CI failure for healing tests."
