import pytest
import settings
from update_facebook import get_data
def test_check():
    assert True
# from update_facebook import get_data
def test_get_data():
    competencies =get_data(settings.INPUT_EXCEL)
    assert competencies.columns.to_list() == ['No', 'Initials', 'Name', '_sorting column', 'Position', 'Delivery Org 1', 'Delivery Org 2', 'Education', 'Experience', 'Professional Competencies', '_professional_competencies', 'Professional Interest']