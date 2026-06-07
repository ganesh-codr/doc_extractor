from doc_extractor.extractor import paginate_rows


def test_paginate_rows_returns_first_page():
    rows = list(range(30))

    result = paginate_rows(rows, page_number=1, page_size=25)

    assert result['page'] == 1
    assert result['page_count'] == 2
    assert result['total_rows'] == 30
    assert result['rows'] == list(range(25))


def test_paginate_rows_returns_last_page():
    rows = list(range(30))

    result = paginate_rows(rows, page_number=2, page_size=25)

    assert result['page'] == 2
    assert result['rows'] == list(range(25, 30))
