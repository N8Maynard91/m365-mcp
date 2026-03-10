def test_graph_base_url_format():
    """Base URL should end with slash for consistent path concatenation."""
    from m365_mcp_server import GRAPH_BASE_URL

    assert GRAPH_BASE_URL.endswith("/"), "Base URL must end with / for clean path joining"
