import pytest
from restaurant_bill import bill_total, split_bill, table_wait

# ==========================================
# Tests for bill_total
# ==========================================

def test_dine_in_bill():
    """Test standard dine-in bill without happy hour and normal party size."""
    assert bill_total("dine_in", 1000, 2, False) == 1155

def test_dine_in_happy_hour():
    """Test dine-in bill during happy hour (20% discount on food)."""
    # food = 1000 * 0.8 = 800, service = 80, total = 880, tax = 44 => 924
    assert bill_total("dine_in", 1000, 2, True) == 924

def test_large_party_service_surcharge():
    """Test party of 8 or more diners gets 200 service surcharge."""
    # food = 1000, service = 100 + 200 = 300, total = 1300, tax = 65 => 1365
    assert bill_total("dine_in", 1000, 8, False) == 1365

def test_large_party_happy_hour():
    """Test party of 8 or more diners with happy hour."""
    # food = 800, service = 80 + 200 = 280, total = 1080, tax = 54 => 1134
    assert bill_total("dine_in", 1000, 8, True) == 1134

def test_takeaway_happy_hour_no_discount():
    """Test takeaway section: 0 service charge and happy hour does not discount takeaway."""
    # food = 1000, service = 0, total = 1000, tax = 50 => 1050
    assert bill_total("takeaway", 1000, 2, True) == 1050

def test_takeaway_regular():
    """Test takeaway section without happy hour."""
    assert bill_total("takeaway", 1000, 2, False) == 1050

def test_rooftop_service_charge():
    """Test rooftop section with 15% service charge."""
    # food = 1000, service = 150, total = 1150, tax = 57 => 1207
    assert bill_total("rooftop", 1000, 4, False) == 1207

def test_rooftop_happy_hour():
    """Test rooftop section with happy hour."""
    # food = 800, service = 120, total = 920, tax = 46 => 966
    assert bill_total("rooftop", 1000, 4, True) == 966

def test_bill_total_unknown_section():
    """Test exception when invalid seating section is provided."""
    with pytest.raises(ValueError, match="unknown section"):
        bill_total("balcony", 1000, 2, False)

def test_bill_total_negative_food_total():
    """Test exception when food_total is negative."""
    with pytest.raises(ValueError, match="invalid order"):
        bill_total("dine_in", -50, 2, False)

def test_bill_total_invalid_guests():
    """Test exception when guest count is less than 1."""
    with pytest.raises(ValueError, match="invalid order"):
        bill_total("dine_in", 1000, 0, False)


# ==========================================
# Tests for split_bill
# ==========================================

def test_even_split():
    """Test even split with tip."""
    # 1000 + 10% = 1100 // 4 = 275 (exact)
    assert split_bill(1000, 4, 10) == 275

def test_uneven_split_rounds_up():
    """Test uneven split rounds up share when there is a remainder."""
    # 100 + 0% = 100 // 3 = 33, remainder 1 != 0 => share = 34
    assert split_bill(100, 3, 0) == 34

def test_split_bill_invalid_people():
    """Test exception when people count is less than 1."""
    with pytest.raises(ValueError, match="invalid people count"):
        split_bill(1000, 0, 10)

def test_split_bill_negative_tip():
    """Test exception when tip percent is negative."""
    with pytest.raises(ValueError, match="invalid tip"):
        split_bill(1000, 4, -5)


# ==========================================
# Tests for table_wait
# ==========================================

def test_table_wait_large_party_calls_ahead():
    """Party size greater than 12 must call ahead."""
    assert table_wait(13, 5, True) == "call ahead"
    assert table_wait(14, 0, False) == "call ahead"

def test_table_wait_with_reservation():
    """Party with reservation is seated now."""
    assert table_wait(4, 0, True) == "seat now"
    assert table_wait(8, 0, True) == "seat now"

def test_table_wait_no_free_tables():
    """Party without reservation and no free tables is waitlisted."""
    assert table_wait(4, 0, False) == "waitlist"

def test_table_wait_party_over_6_few_tables():
    """Party size > 6 with fewer than 2 tables is waitlisted."""
    assert table_wait(7, 1, False) == "waitlist"

def test_table_wait_party_over_6_sufficient_tables():
    """Party size > 6 with 2 or more tables is seated now."""
    assert table_wait(7, 2, False) == "seat now"

def test_table_wait_small_party_one_table():
    """Party size <= 6 with at least 1 table is seated now."""
    assert table_wait(4, 1, False) == "seat now"
    assert table_wait(6, 1, False) == "seat now"
