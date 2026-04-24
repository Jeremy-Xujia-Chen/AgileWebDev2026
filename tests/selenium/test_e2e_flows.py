"""Browser E2E tests: require Chrome/Chromium; Selenium Manager resolves the driver (Selenium 4.6+)."""

from __future__ import annotations

import time
import uuid

import pytest
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

pytestmark = pytest.mark.e2e


def _wait(driver, s=15):
    return WebDriverWait(driver, s)


def _js_click(driver, el) -> None:
    """Avoids ElementClickIntercepted (footer / overlay) in small headless viewports."""
    driver.execute_script("arguments[0].click();", el)


def _login(driver, base: str, email: str, password: str) -> None:
    driver.get(f"{base}/login")
    _wait(driver).until(EC.presence_of_element_located((By.ID, "login-email")))
    driver.find_element(By.ID, "login-email").clear()
    driver.find_element(By.ID, "login-email").send_keys(email)
    driver.find_element(By.ID, "login-password").send_keys(password)
    btn = driver.find_element(By.CSS_SELECTOR, "#loginForm input.btn-primary-custom")
    _js_click(driver, btn)
    _wait(driver, 20).until(EC.url_contains("/timetable"))


def test_e2e_health_reachable(e2e_base_url, e2e_driver) -> None:
    e2e_driver.get(f"{e2e_base_url}/health")
    src = e2e_driver.page_source.lower()
    assert "ok" in src and "studysync" in src


def test_e2e_register_lands_on_timetable(e2e_base_url, e2e_driver) -> None:
    uid = uuid.uuid4().hex[:8]
    email = f"reg_{uid}@student.uwa.edu.au"
    e2e_driver.get(f"{e2e_base_url}/login?tab=register")
    _wait(e2e_driver).until(EC.visibility_of_element_located((By.ID, "reg-email")))
    e2e_driver.find_element(By.ID, "reg-full_name").send_keys("E2E Register")
    e2e_driver.find_element(By.ID, "reg-email").send_keys(email)
    e2e_driver.find_element(By.ID, "reg-password").send_keys("E2E_pass_99")
    _js_click(e2e_driver, e2e_driver.find_element(By.CSS_SELECTOR, "#registerForm input.btn-primary-custom"))
    _wait(e2e_driver, 20).until(EC.url_contains("/timetable"))
    title = e2e_driver.find_element(By.CSS_SELECTOR, ".page-title").text.strip()
    assert "Timetable" in title


def test_e2e_login_shows_timetable(e2e_user, e2e_base_url, e2e_driver) -> None:
    _login(e2e_driver, e2e_base_url, e2e_user["email"], e2e_user["password"])
    title = e2e_driver.find_element(By.CSS_SELECTOR, ".page-title").text
    assert "Timetable" in title

    e2e_driver.get(f"{e2e_base_url}/api/auth/me")
    assert e2e_user["email"] in e2e_driver.page_source
    assert "authenticated" in e2e_driver.page_source


def test_e2e_timetable_create_event_modal(e2e_user, e2e_base_url, e2e_driver) -> None:
    """Create form is pre-filled with a slot on the current week (JS default); we only set title and save."""
    _login(e2e_driver, e2e_base_url, e2e_user["email"], e2e_user["password"])
    e2e_driver.find_element(By.ID, "btnCreateEvent").click()
    _wait(e2e_driver).until(EC.visibility_of_element_located((By.ID, "eventModal")))
    e2e_driver.find_element(By.ID, "evTitle").send_keys("E2E Lecture")
    _js_click(e2e_driver, e2e_driver.find_element(By.ID, "btnSaveEvent"))
    time.sleep(0.7)
    err = e2e_driver.find_element(By.ID, "loadError")
    assert "d-none" in (err.get_attribute("class") or ""), e2e_driver.find_element(
        By.TAG_NAME, "body"
    ).text
    e2e_driver.get(f"{e2e_base_url}/timetable")
    time.sleep(0.8)
    assert "E2E" in e2e_driver.find_element(By.TAG_NAME, "body").text


def test_e2e_exams_list_nav(e2e_user, e2e_base_url, e2e_driver) -> None:
    _login(e2e_driver, e2e_base_url, e2e_user["email"], e2e_user["password"])
    link = e2e_driver.find_element(
        By.XPATH, "//a[contains(@href,'/exams') and contains(.,'Exams')]"
    )
    _js_click(e2e_driver, link)
    _wait(e2e_driver, 20).until(EC.url_contains("/exams"))
    head = e2e_driver.find_element(By.CSS_SELECTOR, ".page-title").text
    assert "Exams" in head
    e2e_driver.find_element(By.ID, "btnNewExam")  # modal control present
    e2e_driver.get(f"{e2e_base_url}/api/exams/sessions")
    assert "sessions" in e2e_driver.page_source


def test_e2e_group_create_flow(e2e_user, e2e_base_url, e2e_driver) -> None:
    _login(e2e_driver, e2e_base_url, e2e_user["email"], e2e_user["password"])
    e2e_driver.get(f"{e2e_base_url}/group")
    _wait(e2e_driver).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-bs-target='#modalCreateGroup']")))
    e2e_driver.find_element(
        By.CSS_SELECTOR, "button[data-bs-target='#modalCreateGroup']"
    ).click()
    gname = f"E2EGroup_{uuid.uuid4().hex[:6]}"
    _wait(e2e_driver).until(EC.visibility_of_element_located((By.ID, "inputCreateName")))
    e2e_driver.find_element(By.ID, "inputCreateName").send_keys(gname)
    _js_click(e2e_driver, e2e_driver.find_element(By.ID, "btnCreateGroupSubmit"))
    _wait(e2e_driver, 20).until(EC.text_to_be_present_in_element((By.ID, "dispGroupName"), gname))
    _wait(e2e_driver, 5).until(
        lambda d: "d-none" not in (d.find_element(By.ID, "panelGroup").get_attribute("class") or "")
    )
    code = e2e_driver.find_element(By.ID, "dispJoinCode").text
    assert code.strip() and len(code.strip()) >= 6


def test_e2e_logout_returns_to_login(e2e_user, e2e_base_url, e2e_driver) -> None:
    _login(e2e_driver, e2e_base_url, e2e_user["email"], e2e_user["password"])
    _js_click(
        e2e_driver, e2e_driver.find_element(By.CSS_SELECTOR, "form[action$='logout'] button")
    )
    _wait(e2e_driver, 20).until(EC.url_contains("/login"))
    e2e_driver.get(f"{e2e_base_url}/timetable")
    time.sleep(0.3)
    # Redirect to login
    _wait(e2e_driver, 20).until(EC.url_contains("/login"))
