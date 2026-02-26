"""classifiers モジュールのテスト."""

from classifiers import classify_document
from models import TextElement


def _make_elements(words: list[str]) -> list[TextElement]:
    """y1 が降順になるようなダミー要素を生成する."""
    elements = []
    for i, word in enumerate(words):
        y1 = 800.0 - i * 10
        elements.append(TextElement(word=word, x0=0.0, y0=y1 - 10, x1=100.0, y1=y1))
    return elements


def test_classify_final_check():
    elements = _make_elements(["header", "sub", "ユニットバスルーム納期最終確認票", "body"])
    assert classify_document(elements) == "final_check"


def test_classify_final_check_at_index_1():
    elements = _make_elements(["header", "ユニットバスルーム納期最終確認票", "body", "body2"])
    assert classify_document(elements) == "final_check"


def test_classify_quotation():
    elements = _make_elements(["御 見 積 書", "sub", "body", "body2"])
    assert classify_document(elements) == "quotation"


def test_classify_detail():
    elements = _make_elements(["header", "ユニットバスルームご発注確認票", "body", "body2"])
    assert classify_document(elements) == "detail"


def test_classify_change_spec():
    elements = _make_elements(["header", "仕様変更確認票", "body", "body2"])
    assert classify_document(elements) == "change_spec"


def test_classify_cancel():
    elements = _make_elements(["キ ャ ン セ ル 確 認 票", "sub", "body", "body2"])
    assert classify_document(elements) == "cancel"


def test_classify_unknown():
    elements = _make_elements(["unknown", "document", "type", "here"])
    assert classify_document(elements) is None
