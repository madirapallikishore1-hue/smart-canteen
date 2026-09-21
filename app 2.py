
import streamlit as st
import pandas as pd
import os
import json
import uuid
import secrets
import hashlib
from datetime import datetime, timedelta
from io import BytesIO

# Optional libraries used only for report downloads
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    )
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False


# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Smart Canteen Queue & Order Management",
    page_icon="🍔",
    layout="wide",
    initial_sidebar_state="expanded"
)

DATA_DIR = "data"
LOGIN_SESSIONS_FILE = os.path.join(DATA_DIR, "login_sessions.json")
USERS_FILE = os.path.join(DATA_DIR, "users.csv")
ORDERS_FILE = os.path.join(DATA_DIR, "orders.csv")
PAYMENTS_FILE = os.path.join(DATA_DIR, "payments.csv")
OTP_FILE = os.path.join(DATA_DIR, "otp_sessions.json")

# Persistent login: 10 years. It is intentionally removed immediately on Logout.
PERSISTENT_LOGIN_DAYS = 3650
OTP_VALID_MINUTES = 5

# Optional real SMS provider (Twilio). Set these environment variables for real SMS.
try:
    from twilio.rest import Client as TwilioClient
    TWILIO_OK = True
except ImportError:
    TwilioClient = None
    TWILIO_OK = False

os.makedirs(DATA_DIR, exist_ok=True)

FOODS = [
    {"Food": "Veg Burger", "Emoji": "🍔", "Price": 80, "Category": "Fast Food"},
    {"Food": "Pizza", "Emoji": "🍕", "Price": 140, "Category": "Fast Food"},
    {"Food": "Sandwich", "Emoji": "🥪", "Price": 60, "Category": "Fast Food"},
    {"Food": "French Fries", "Emoji": "🍟", "Price": 70, "Category": "Snacks"},
    {"Food": "Veg Biryani", "Emoji": "🍚", "Price": 120, "Category": "Meals"},
    {"Food": "Masala Dosa", "Emoji": "🥞", "Price": 70, "Category": "South Indian"},
    {"Food": "Idly", "Emoji": "🍘", "Price": 40, "Category": "South Indian"},
    {"Food": "Tea", "Emoji": "☕", "Price": 20, "Category": "Beverages"},
]

DEFAULT_USERS = pd.DataFrame(
    [
        ["STU001", "Kishore", "kishore@gmail.com", "123456", "9876543210"],
        ["STU002", "Ravi", "ravi@gmail.com", "123456", "9876543211"],
    ],
    columns=["Student_ID", "Name", "Email", "Password", "Phone"]
)

USER_COLUMNS = ["Student_ID", "Name", "Email", "Password", "Phone"]

ORDER_COLUMNS = [
    "Order_ID", "Student_ID", "Student_Name", "Food_Items",
    "Total_Amount", "Order_Date", "Payment_Status", "Order_Status",
    "Token_No", "Queue_Position", "Estimated_Wait_Minutes",
    "Pickup_Counter"
]

PAYMENT_COLUMNS = [
    "Payment_ID", "Order_ID", "Student_ID", "Amount",
    "Payment_Method", "Payment_Date", "Payment_Status", "Transaction_ID"
]

# Demo owner account used only for this project application.
# These credentials are NOT displayed on the login screen.
OWNER_EMAIL = "owner@canteen.com"
OWNER_PASSWORD = "owner123"
# Registered canteen owner mobile number used for OTP verification.
# Change this to the real owner mobile number before deployment.
OWNER_PHONE = os.getenv("OWNER_PHONE", "9876543212")


# ============================================================
# FILE / DATA HELPERS
# ============================================================

def load_login_sessions():
    """Load persistent login sessions stored on the server."""
    if not os.path.exists(LOGIN_SESSIONS_FILE):
        return {}
    try:
        with open(LOGIN_SESSIONS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_login_sessions(sessions):
    """Save persistent login sessions safely."""
    os.makedirs(DATA_DIR, exist_ok=True)
    temp_file = LOGIN_SESSIONS_FILE + ".tmp"
    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(sessions, f, indent=2)
    os.replace(temp_file, LOGIN_SESSIONS_FILE)


def create_login_session(role, student_id="", student_name=""):
    """Create a long-lived server-side session and return its random token."""
    token = uuid.uuid4().hex + uuid.uuid4().hex
    sessions = load_login_sessions()
    now = datetime.now()
    sessions[token] = {
        "role": role,
        "student_id": str(student_id or ""),
        "student_name": str(student_name or ""),
        "created_at": now.isoformat(),
        "last_seen": now.isoformat(),
        "expires_at": (now + timedelta(days=PERSISTENT_LOGIN_DAYS)).isoformat(),
    }
    save_login_sessions(sessions)
    return token


def get_login_session(token):
    """Validate a persistent login token and return its session data."""
    if not token:
        return None
    sessions = load_login_sessions()
    session = sessions.get(str(token))
    if not session:
        return None
    try:
        expires_at = datetime.fromisoformat(session.get("expires_at", ""))
    except Exception:
        return None
    if datetime.now() >= expires_at:
        sessions.pop(str(token), None)
        save_login_sessions(sessions)
        return None
    session["last_seen"] = datetime.now().isoformat()
    sessions[str(token)] = session
    save_login_sessions(sessions)
    return session


def delete_login_session(token):
    if not token:
        return
    sessions = load_login_sessions()
    if str(token) in sessions:
        sessions.pop(str(token), None)
        save_login_sessions(sessions)


def load_otp_sessions():
    if not os.path.exists(OTP_FILE):
        return {}
    try:
        with open(OTP_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_otp_sessions(data):
    with open(OTP_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def normalize_phone(phone):
    digits = "".join(ch for ch in str(phone) if ch.isdigit())
    if digits.startswith("91") and len(digits) == 12:
        return "+" + digits
    if len(digits) == 10:
        return "+91" + digits
    if str(phone).strip().startswith("+") and len(digits) >= 10:
        return "+" + digits
    return ""


def send_sms_otp(phone, otp):
    sid = os.getenv("TWILIO_ACCOUNT_SID", "").strip()
    auth = os.getenv("TWILIO_AUTH_TOKEN", "").strip()
    from_number = os.getenv("TWILIO_FROM_NUMBER", "").strip()
    if not (TWILIO_OK and sid and auth and from_number):
        return False, "SMS provider is not configured."
    try:
        client = TwilioClient(sid, auth)
        client.messages.create(
            body=f"Smart Canteen verification OTP: {otp}. Valid for {OTP_VALID_MINUTES} minutes.",
            from_=from_number,
            to=phone
        )
        return True, "OTP sent to your mobile number."
    except Exception as exc:
        return False, f"SMS could not be sent: {exc}"


def create_otp(phone):
    """Create an OTP, save only its hash, and send it by SMS when configured.

    Returns exactly two values: (success, message).
    This fixes the 'too many values to unpack' error in login_page().
    """
    otp = f"{secrets.randbelow(1000000):06d}"

    sessions = load_otp_sessions()
    now = datetime.now()
    sessions[phone] = {
        "otp_hash": hashlib.sha256(otp.encode()).hexdigest(),
        "created_at": now.isoformat(),
        "expires_at": (now + timedelta(minutes=OTP_VALID_MINUTES)).isoformat(),
        "attempts": 0
    }
    save_otp_sessions(sessions)

    # Keep the OTP only in Streamlit session state for development testing.
    # It is not written to the OTP JSON file in plain text.
    st.session_state.otp_dev_code = ""

    sms_ok, sms_message = send_sms_otp(phone, otp)

    if sms_ok:
        return True, sms_message

    # Development fallback: if Twilio is not configured, show the OTP
    # on the page so the project can still be tested locally.
    st.session_state.otp_dev_code = otp

    if sms_message == "SMS provider is not configured.":
        return True, "Development OTP generated. Configure Twilio for real SMS."

    return True, f"Development OTP generated because SMS could not be sent: {sms_message}"


def verify_otp(phone, entered_otp):
    sessions = load_otp_sessions()
    record = sessions.get(phone)
    if not record:
        return False, "Please request a new OTP."
    try:
        expires_at = datetime.fromisoformat(record.get("expires_at", ""))
    except Exception:
        sessions.pop(phone, None)
        save_otp_sessions(sessions)
        return False, "OTP expired. Please request a new OTP."
    if datetime.now() >= expires_at:
        sessions.pop(phone, None)
        save_otp_sessions(sessions)
        return False, "OTP expired. Please request a new OTP."
    if int(record.get("attempts", 0)) >= 5:
        sessions.pop(phone, None)
        save_otp_sessions(sessions)
        return False, "Too many attempts. Please request a new OTP."
    record["attempts"] = int(record.get("attempts", 0)) + 1
    sessions[phone] = record
    save_otp_sessions(sessions)
    entered_hash = hashlib.sha256(str(entered_otp).strip().encode()).hexdigest()
    if secrets.compare_digest(entered_hash, record.get("otp_hash", "")):
        sessions.pop(phone, None)
        save_otp_sessions(sessions)
        return True, "Mobile number verified successfully."
    return False, "Invalid OTP."


def ensure_csv_files():
    if not os.path.exists(USERS_FILE):
        DEFAULT_USERS.to_csv(USERS_FILE, index=False)

    if not os.path.exists(ORDERS_FILE):
        pd.DataFrame(columns=ORDER_COLUMNS).to_csv(ORDERS_FILE, index=False)

    if not os.path.exists(PAYMENTS_FILE):
        pd.DataFrame(columns=PAYMENT_COLUMNS).to_csv(PAYMENTS_FILE, index=False)


def load_users():
    ensure_csv_files()
    try:
        df = pd.read_csv(USERS_FILE, dtype=str).fillna("")
    except Exception:
        df = DEFAULT_USERS.copy()
        df.to_csv(USERS_FILE, index=False)

    for col in USER_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[USER_COLUMNS]


def save_users(df):
    df[USER_COLUMNS].to_csv(USERS_FILE, index=False)


def load_orders():
    ensure_csv_files()
    try:
        df = pd.read_csv(ORDERS_FILE, dtype=str).fillna("")
    except Exception:
        df = pd.DataFrame(columns=ORDER_COLUMNS)

    for col in ORDER_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[ORDER_COLUMNS]


def save_orders(df):
    df[ORDER_COLUMNS].to_csv(ORDERS_FILE, index=False)


def load_payments():
    ensure_csv_files()
    try:
        df = pd.read_csv(PAYMENTS_FILE, dtype=str).fillna("")
    except Exception:
        df = pd.DataFrame(columns=PAYMENT_COLUMNS)

    for col in PAYMENT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[PAYMENT_COLUMNS]


def save_payments(df):
    df[PAYMENT_COLUMNS].to_csv(PAYMENTS_FILE, index=False)


def money(value):
    try:
        return f"₹{float(value):,.0f}"
    except Exception:
        return "₹0"


def next_number(df, column, prefix, width=3):
    if df.empty or column not in df.columns:
        return f"{prefix}{1:0{width}d}"

    nums = []
    for value in df[column].astype(str):
        digits = "".join(ch for ch in value if ch.isdigit())
        if digits:
            try:
                nums.append(int(digits))
            except Exception:
                pass

    number = max(nums, default=0) + 1
    return f"{prefix}{number:0{width}d}"


def parse_items(text):
    if not text:
        return []
    try:
        data = json.loads(text)
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []


def cart_total(cart):
    return sum(item["Price"] * item["Quantity"] for item in cart.values())


def cart_count(cart):
    return sum(item["Quantity"] for item in cart.values())


def reset_to(page):
    st.session_state.page = page
    st.rerun()


# ============================================================
# SESSION STATE
# ============================================================

def init_state():
    defaults = {
        "logged_in": False,
        "user_role": "",
        "student_id": "",
        "student_name": "",
        "page": "Dashboard",
        "cart": {},
        "checkout_order": None,
        "last_payment": None,
        "last_order": None,
        "persistent_session_token": "",
        "otp_phone": "",
        "otp_sent": False,
        "otp_dev_code": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_state()
ensure_csv_files()

# Persistent browser login. The cookie stores only a random session token;
# the actual account details remain on the server in login_sessions.json.
try:
    from streamlit_cookies_controller import CookieController
    cookie_controller = CookieController(key="smart_canteen_cookie_controller")
    COOKIE_OK = True
except ImportError:
    cookie_controller = None
    COOKIE_OK = False


def restore_persistent_login():
    if st.session_state.get("logged_in") or not COOKIE_OK:
        return
    try:
        token = cookie_controller.get("smart_canteen_session")
        session = get_login_session(token)
        if session:
            st.session_state.logged_in = True
            st.session_state.user_role = session.get("role", "")
            st.session_state.student_id = session.get("student_id", "")
            st.session_state.student_name = session.get("student_name", "")
            st.session_state.page = (
                "Owner Dashboard" if session.get("role") == "owner" else "Dashboard"
            )
            st.session_state.persistent_session_token = token
    except Exception:
        pass


def start_persistent_login(role, student_id="", student_name=""):
    if not COOKIE_OK:
        return
    try:
        token = create_login_session(role, student_id, student_name)
        cookie_controller.set(
            "smart_canteen_session",
            token,
            max_age=PERSISTENT_LOGIN_DAYS * 24 * 60 * 60,
            path="/"
        )
        st.session_state.persistent_session_token = token
    except Exception:
        pass


restore_persistent_login()


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>
    .main-title {
        font-size: 42px;
        font-weight: 800;
        margin-bottom: 0;
    }
    .sub-title {
        font-size: 20px;
        color: #64748b;
        margin-top: 4px;
    }
    .food-card {
        border: 1px solid #e5e7eb;
        border-radius: 18px;
        padding: 18px;
        margin-bottom: 16px;
        background: #ffffff;
        box-shadow: 0 4px 14px rgba(0,0,0,0.06);
    }
    .food-name {
        font-size: 22px;
        font-weight: 700;
    }
    .food-price {
        font-size: 21px;
        font-weight: 700;
    }
    .success-box {
        padding: 20px;
        border-radius: 18px;
        background: #ecfdf5;
        border: 1px solid #86efac;
    }
    .token-box {
        padding: 28px;
        border-radius: 20px;
        background: #fff7ed;
        border: 2px dashed #fb923c;
        text-align: center;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# LOGIN / REGISTRATION
# ============================================================

def login_page():
    st.markdown(
        '<div class="main-title">🍔 Smart Canteen Queue & Order Management</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sub-title">Student & Canteen Owner Login</div>',
        unsafe_allow_html=True
    )
    st.write("")

    student_tab, owner_tab, register_tab = st.tabs(
        ["🎓 Student Login", "👨‍🍳 Owner Login", "📝 Registration"]
    )

    # ---------------- STUDENT LOGIN ----------------
    with student_tab:
        st.info("🔐 First login/after Logout: verify your registered mobile number with OTP. After that, this browser stays signed in until you press Logout.")

        phone = st.text_input(
            "Registered Mobile Number",
            placeholder="Enter your 10-digit mobile number",
            key="student_phone_login"
        )
        normalized = normalize_phone(phone)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("📱 Send OTP", type="primary", use_container_width=True):
                users = load_users()
                matched = users[users["Phone"].astype(str).str.replace(r"\D", "", regex=True) == normalized.replace("+91", "")]
                if not normalized or len(normalized.replace("+91", "")) != 10:
                    st.error("Enter a valid 10-digit Indian mobile number.")
                elif matched.empty:
                    st.error("This mobile number is not registered. Please register first.")
                else:
                    # create_otp() returns (success, message) and stores the
                    # development OTP in session_state when real SMS is unavailable.
                    ok, message = create_otp(normalized)
                    st.session_state.otp_phone = normalized
                    st.session_state.otp_sent = bool(ok)

                    if ok:
                        # Real SMS was sent, or development mode generated an OTP.
                        if TWILIO_OK and all([
                            os.getenv("TWILIO_ACCOUNT_SID"),
                            os.getenv("TWILIO_AUTH_TOKEN"),
                            os.getenv("TWILIO_FROM_NUMBER")
                        ]):
                            st.success("OTP sent to your registered mobile number.")
                        else:
                            st.warning("SMS service is not configured. Development OTP is shown below. For real mobile SMS, configure Twilio.")
                            dev_code = st.session_state.get("otp_dev_code", "")
                            if dev_code:
                                st.code(dev_code, language="text")
                        # Do not display the raw tuple returned by create_otp().
                        if message and "Development OTP" not in message and not (
                            TWILIO_OK and all([
                                os.getenv("TWILIO_ACCOUNT_SID"),
                                os.getenv("TWILIO_AUTH_TOKEN"),
                                os.getenv("TWILIO_FROM_NUMBER")
                            ])
                        ):
                            st.caption(message)
                    else:
                        st.session_state.otp_sent = False
                        st.error(message)

        with c2:
            if st.button("🔄 Change Number", use_container_width=True):
                st.session_state.otp_phone = ""
                st.session_state.otp_sent = False
                st.session_state.otp_dev_code = ""
                st.rerun()

        if st.session_state.get("otp_sent") and st.session_state.get("otp_phone"):
            st.write("")
            otp = st.text_input(
                "Enter 6-digit OTP",
                max_chars=6,
                placeholder="Enter OTP received on your mobile",
                key="student_otp"
            )
            if st.button("✅ Verify Mobile & Login", type="primary", use_container_width=True):
                valid, message = verify_otp(st.session_state.otp_phone, otp)
                if valid:
                    users = load_users()
                    phone_digits = st.session_state.otp_phone.replace("+91", "")
                    match = users[users["Phone"].astype(str).str.replace(r"\D", "", regex=True) == phone_digits]
                    if not match.empty:
                        user = match.iloc[0]
                        st.session_state.logged_in = True
                        st.session_state.user_role = "student"
                        st.session_state.student_id = str(user["Student_ID"])
                        st.session_state.student_name = str(user["Name"])
                        st.session_state.page = "Dashboard"
                        st.session_state.otp_sent = False
                        st.session_state.otp_dev_code = ""
                        start_persistent_login("student", user["Student_ID"], user["Name"])
                        st.success("Mobile verified. Login successful! You will stay logged in until Logout.")
                        st.rerun()
                    else:
                        st.error("Account not found for this mobile number.")
                else:
                    st.error(message)

    # ---------------- OWNER LOGIN ----------------
    # Owner login uses mobile-number OTP.
    # After successful OTP verification, the owner remains logged in
    # until the Logout button is used.
    with owner_tab:
        st.info("Canteen Owner Login: verify the registered mobile number with OTP.")

        owner_phone_input = st.text_input(
            "Owner Mobile Number",
            placeholder="Enter registered owner mobile number",
            key="owner_phone_input"
        )

        owner_send_col, owner_change_col = st.columns(2)
        with owner_send_col:
            owner_send_otp = st.button(
                "📲 Send Owner OTP",
                key="owner_send_otp",
                type="primary"
            )
        with owner_change_col:
            owner_change_number = st.button(
                "🔄 Change Number",
                key="owner_change_number"
            )

        if owner_change_number:
            st.session_state.otp_phone = ""
            st.session_state.otp_sent = False
            st.session_state.otp_dev_code = ""
            st.rerun()

        if owner_send_otp:
            normalized_owner_phone = normalize_phone(owner_phone_input)
            if normalized_owner_phone != normalize_phone(OWNER_PHONE):
                st.error("This mobile number is not registered for the canteen owner.")
            else:
                success, message = create_otp(normalized_owner_phone)
                if success:
                    st.session_state.otp_phone = normalized_owner_phone
                    st.session_state.otp_sent = True
                    st.success("OTP sent to the registered owner mobile number.")
                    # Development fallback when SMS provider is not configured.
                    if not TWILIO_OK or not all([
                        os.getenv("TWILIO_ACCOUNT_SID"),
                        os.getenv("TWILIO_AUTH_TOKEN"),
                        os.getenv("TWILIO_FROM_NUMBER")
                    ]):
                        st.warning(
                            f"Development OTP: {st.session_state.get('otp_dev_code', '')}"
                        )
                else:
                    st.error(message)

        if st.session_state.get("otp_sent") and st.session_state.get("otp_phone") == normalize_phone(owner_phone_input):
            owner_otp = st.text_input(
                "Enter 6-Digit Owner OTP",
                max_chars=6,
                type="password",
                placeholder="Enter OTP",
                key="owner_otp_input"
            )
            owner_verify = st.button(
                "🔐 Verify Owner Mobile & Login",
                key="owner_verify_otp",
                type="primary"
            )

            if owner_verify:
                valid, message = verify_otp(
                    normalize_phone(owner_phone_input),
                    owner_otp
                )
                if valid:
                    st.session_state.logged_in = True
                    st.session_state.user_role = "owner"
                    st.session_state.student_id = ""
                    st.session_state.student_name = "Canteen Owner"
                    st.session_state.page = "Owner Dashboard"
                    start_persistent_login("owner", "", "Canteen Owner")
                    st.session_state.otp_phone = ""
                    st.session_state.otp_sent = False
                    st.session_state.otp_dev_code = ""
                    st.success("Owner mobile verified. Login successful! You will stay logged in until Logout.")
                    st.rerun()
                else:
                    st.error(message)

    # ---------------- STUDENT REGISTRATION ----------------
    with register_tab:
        with st.form("register_form"):
            name = st.text_input("Full Name")
            email = st.text_input("Email Address")
            password = st.text_input("Password", type="password")
            confirm = st.text_input("Confirm Password", type="password")
            phone = st.text_input("Phone Number")
            submitted = st.form_submit_button(
                "Create Student Account",
                type="primary"
            )

            if submitted:
                users = load_users()
                email_clean = email.strip().lower()

                if not name.strip() or not email_clean or not password or not phone.strip():
                    st.error("Please fill all fields.")
                elif password != confirm:
                    st.error("Passwords do not match.")
                elif users["Email"].astype(str).str.lower().eq(email_clean).any():
                    st.error("Email already registered.")
                else:
                    student_id = next_number(users, "Student_ID", "STU", 3)
                    new_user = pd.DataFrame(
                        [[student_id, name.strip(), email_clean, password, phone.strip()]],
                        columns=USER_COLUMNS
                    )
                    users = pd.concat([users, new_user], ignore_index=True)
                    save_users(users)
                    st.success(
                        f"Registration successful! Your Student ID is {student_id}."
                    )
                    st.info("Go to the Student Login tab and verify your registered mobile number with OTP.")


# ============================================================
# SIDEBAR
# ============================================================

def sidebar():
    with st.sidebar:
        st.markdown("## 🍔 Smart Canteen")
        st.caption("Queue & Order Management")
        st.divider()

        if st.session_state.user_role == "owner":
            st.markdown("### 👨‍🍳 Canteen Owner")
            st.caption("Owner Control Panel")

            menu = [
                ("🏠 Owner Dashboard", "Owner Dashboard"),
                ("📦 Received Orders", "Received Orders"),
                ("📊 Annual Report", "Annual Report"),
            ]

            for label, page in menu:
                if st.button(
                    label,
                    key=f"owner_side_{page}",
                    use_container_width=True,
                    type="primary" if st.session_state.page == page else "secondary"
                ):
                    st.session_state.page = page
                    st.rerun()

        else:
            st.markdown(f"### 👤 {st.session_state.student_name}")
            st.caption(f"Student ID: {st.session_state.student_id}")

            menu = [
                ("🏠 Dashboard", "Dashboard"),
                ("🍔 Food Menu", "Food Menu"),
                ("🛒 Cart", "Cart"),
                ("💳 Checkout", "Checkout"),
                ("💰 Payment", "Payment"),
                ("🎟️ Token", "Token"),
                ("🧍 Queue", "Queue"),
                ("📋 Order Details", "Order Details"),
                ("📊 Annual Report", "Annual Report"),
            ]

            for label, page in menu:
                if st.button(
                    label,
                    key=f"side_{page}",
                    use_container_width=True,
                    type="primary" if st.session_state.page == page else "secondary"
                ):
                    st.session_state.page = page
                    st.rerun()

        st.divider()

        if st.button("🚪 Logout", use_container_width=True):
            # Explicit logout removes both the server session and browser cookie.
            delete_login_session(st.session_state.get("persistent_session_token", ""))
            if COOKIE_OK:
                try:
                    cookie_controller.remove("smart_canteen_session")
                except Exception:
                    pass
            st.session_state.logged_in = False
            st.session_state.user_role = ""
            st.session_state.student_id = ""
            st.session_state.student_name = ""
            st.session_state.cart = {}
            st.session_state.checkout_order = None
            st.session_state.last_payment = None
            st.session_state.last_order = None
            st.session_state.page = "Dashboard"
            st.rerun()


# ============================================================
# DASHBOARD
# ============================================================

def dashboard():
    orders = load_orders()
    sid = st.session_state.student_id
    mine = orders[orders["Student_ID"].astype(str) == str(sid)].copy()

    active_status = ["Order Received", "Preparing", "Ready for Pickup"]
    active = mine[mine["Order_Status"].isin(active_status)]
    completed = mine[mine["Order_Status"].isin(["Completed", "Picked Up"])]

    spent = pd.to_numeric(mine["Total_Amount"], errors="coerce").fillna(0).sum()

    st.markdown(
        f'<div class="main-title">Welcome, {st.session_state.student_name}! 👋</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="sub-title">Welcome to your Smart Canteen Student Dashboard</div>',
        unsafe_allow_html=True
    )
    st.write("")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Active Orders", len(active))
    c2.metric("Completed Orders", len(completed))
    c3.metric("Cart Items", cart_count(st.session_state.cart))
    c4.metric("Total Spent", money(spent))

    st.write("")
    st.subheader("⚡ Quick Actions")

    a, b, c = st.columns(3)

    with a:
        st.info("🍔 **Food Menu**\n\nSelect your favourite food.")
        if st.button("Open Food Menu →", key="dash_food"):
            reset_to("Food Menu")

    with b:
        st.info("🛒 **Cart**\n\nReview your selected items.")
        if st.button("Open Cart →", key="dash_cart"):
            reset_to("Cart")

    with c:
        st.info("🎟️ **Token**\n\nPayment first → Token generation.")
        if st.button("View Token →", key="dash_token"):
            reset_to("Token")

    st.write("")
    st.subheader("📌 Student Information")

    p1, p2 = st.columns(2)
    with p1:
        st.markdown("### 👤 Profile")
        st.write(f"**Student ID:** {sid}")
        st.write(f"**Name:** {st.session_state.student_name}")

        users = load_users()
        row = users[users["Student_ID"].astype(str) == str(sid)]
        if not row.empty:
            st.write(f"**Email:** {row.iloc[0]['Email']}")
            st.write(f"**Phone:** {row.iloc[0]['Phone']}")

    with p2:
        st.markdown("### 🍽️ How It Works")
        st.write("1️⃣ Select your food")
        st.write("2️⃣ Add items to Cart")
        st.write("3️⃣ Checkout")
        st.write("4️⃣ Complete Payment")
        st.write("5️⃣ Token is generated after successful payment")
        st.write("6️⃣ Track Queue Position")
        st.write("7️⃣ View Order Details")


# ============================================================
# FOOD MENU
# ============================================================

def food_menu():
    st.title("🍔 Food Menu")
    st.caption("Choose your food and add it to your cart.")

    search = st.text_input("🔎 Search food", placeholder="Veg Burger, Pizza...")
    category = st.selectbox(
        "Category",
        ["All"] + sorted(list(set(item["Category"] for item in FOODS)))
    )

    filtered = FOODS

    if search.strip():
        filtered = [
            x for x in filtered
            if search.lower().strip() in x["Food"].lower()
        ]

    if category != "All":
        filtered = [x for x in filtered if x["Category"] == category]

    if not filtered:
        st.warning("No food found.")
        return

    for i in range(0, len(filtered), 2):
        cols = st.columns(2)

        for j, col in enumerate(cols):
            if i + j >= len(filtered):
                continue

            item = filtered[i + j]

            with col:
                st.markdown(
                    f"""
                    <div class="food-card">
                        <div style="font-size:42px">{item["Emoji"]}</div>
                        <div class="food-name">{item["Food"]}</div>
                        <div style="color:#64748b">{item["Category"]}</div>
                        <div class="food-price">{money(item["Price"])}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                quantity = st.number_input(
                    f"Quantity - {item['Food']}",
                    min_value=1,
                    max_value=10,
                    value=1,
                    step=1,
                    key=f"qty_{item['Food']}"
                )

                if st.button(
                    f"➕ Add {item['Food']} to Cart",
                    key=f"add_{item['Food']}",
                    use_container_width=True
                ):
                    if item["Food"] in st.session_state.cart:
                        st.session_state.cart[item["Food"]]["Quantity"] += quantity
                    else:
                        st.session_state.cart[item["Food"]] = {
                            "Food": item["Food"],
                            "Emoji": item["Emoji"],
                            "Price": item["Price"],
                            "Quantity": quantity,
                        }

                    st.success(f"{quantity} × {item['Food']} added to cart.")


# ============================================================
# CART
# ============================================================

def cart_page():
    st.title("🛒 Cart")

    cart = st.session_state.cart

    if not cart:
        st.info("Your cart is empty.")
        if st.button("🍔 Go to Food Menu"):
            reset_to("Food Menu")
        return

    total = 0

    for food_name, item in list(cart.items()):
        c1, c2, c3, c4 = st.columns([3, 1, 1, 1])

        with c1:
            st.write(f"{item['Emoji']} **{item['Food']}**")
            st.caption(f"{money(item['Price'])} each")

        with c2:
            new_qty = st.number_input(
                "Qty",
                min_value=1,
                max_value=20,
                value=int(item["Quantity"]),
                key=f"cart_qty_{food_name}"
            )
            item["Quantity"] = new_qty

        line_total = item["Price"] * item["Quantity"]
        total += line_total

        with c3:
            st.write(f"**{money(line_total)}**")

        with c4:
            if st.button("🗑️ Remove", key=f"remove_{food_name}"):
                del st.session_state.cart[food_name]
                st.rerun()

        st.divider()

    st.subheader(f"Total: {money(total)}")

    x, y = st.columns(2)

    with x:
        if st.button("🍔 Continue Shopping", use_container_width=True):
            reset_to("Food Menu")

    with y:
        if st.button("💳 Proceed to Checkout", type="primary", use_container_width=True):
            st.session_state.page = "Checkout"
            st.rerun()


# ============================================================
# CHECKOUT
# ============================================================

def checkout_page():
    st.title("💳 Checkout")

    cart = st.session_state.cart

    if not cart:
        st.warning("Your cart is empty. Add food before checkout.")
        if st.button("🍔 Go to Food Menu"):
            reset_to("Food Menu")
        return

    total = cart_total(cart)

    st.subheader("🧾 Order Summary")

    summary = []
    for item in cart.values():
        summary.append({
            "Food": f"{item['Emoji']} {item['Food']}",
            "Quantity": item["Quantity"],
            "Price": money(item["Price"]),
            "Total": money(item["Price"] * item["Quantity"])
        })

    st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

    st.markdown(f"### Total Amount: {money(total)}")

    st.info(
        "Important: Token generation is allowed only after successful payment."
    )

    if st.button("💰 Continue to Payment", type="primary", use_container_width=True):
        st.session_state.checkout_order = {
            "items": list(cart.values()),
            "total": total,
        }
        st.session_state.page = "Payment"
        st.rerun()


# ============================================================
# PAYMENT
# ============================================================

def payment_page():
    st.title("💰 Payment")

    cart = st.session_state.cart
    if not cart:
        st.warning("No items available for payment.")
        return

    total = cart_total(cart)

    st.markdown(f"## Amount to Pay: {money(total)}")

    method = st.radio(
        "Select Payment Method",
        ["UPI", "Debit/Credit Card", "Cash at Counter"]
    )

    st.info(
        "This is a project/demo payment screen. It records a successful "
        "demo transaction; it does not charge a real bank account."
    )

    if method == "UPI":
        upi = st.text_input("UPI ID", placeholder="student@upi")
        if not upi.strip():
            st.caption("Enter a UPI ID for the demo payment.")

    elif method == "Debit/Credit Card":
        card = st.text_input("Card Number", placeholder="1234 5678 9012 3456")
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Expiry", placeholder="MM/YY")
        with c2:
            st.text_input("CVV", type="password", placeholder="123")

    else:
        st.warning("Cash payment will be marked as a demo successful payment.")

    if st.button("✅ Pay Now", type="primary", use_container_width=True):
        # Create order only after the payment button is successfully completed.
        orders = load_orders()
        payments = load_payments()

        order_id = next_number(orders, "Order_ID", "ORD", 3)

        # Queue position = number of currently active paid orders + 1.
        active_statuses = ["Order Received", "Preparing", "Ready for Pickup"]
        active_count = len(orders[orders["Order_Status"].isin(active_statuses)])
        queue_position = active_count + 1

        # Simple estimate for a college-canteen project.
        wait_minutes = queue_position * 5

        # The same token number is stored in the order record.
        # Therefore the student and owner see exactly the same token.
        token_no = next_number(orders, "Token_No", "T-", 3)

        items_for_storage = []
        for item in cart.values():
            items_for_storage.append({
                "Food": item["Food"],
                "Emoji": item["Emoji"],
                "Price": item["Price"],
                "Quantity": item["Quantity"],
            })

        now = datetime.now()
        order_row = {
            "Order_ID": order_id,
            "Student_ID": st.session_state.student_id,
            "Student_Name": st.session_state.student_name,
            "Food_Items": json.dumps(items_for_storage),
            "Total_Amount": total,
            "Order_Date": now.strftime("%Y-%m-%d %H:%M:%S"),
            "Payment_Status": "Paid",
            "Order_Status": "Order Received",
            "Token_No": token_no,
            "Queue_Position": queue_position,
            "Estimated_Wait_Minutes": wait_minutes,
            "Pickup_Counter": "Counter 1",
        }

        orders = pd.concat(
            [orders, pd.DataFrame([order_row])],
            ignore_index=True
        )
        save_orders(orders)

        payment_id = next_number(payments, "Payment_ID", "PAY", 3)
        transaction_id = "TXN-" + uuid.uuid4().hex[:10].upper()

        payment_row = {
            "Payment_ID": payment_id,
            "Order_ID": order_id,
            "Student_ID": st.session_state.student_id,
            "Amount": total,
            "Payment_Method": method,
            "Payment_Date": now.strftime("%Y-%m-%d %H:%M:%S"),
            "Payment_Status": "Successful",
            "Transaction_ID": transaction_id,
        }

        payments = pd.concat(
            [payments, pd.DataFrame([payment_row])],
            ignore_index=True
        )
        save_payments(payments)

        st.session_state.last_payment = payment_row
        st.session_state.last_order = order_row
        st.session_state.cart = {}
        st.session_state.checkout_order = None

        st.success("Payment Successful! 🎉")
        st.balloons()

        st.session_state.page = "Token"
        st.rerun()


# ============================================================
# TOKEN
# ============================================================

def token_page():
    st.title("🎟️ Token")

    order = st.session_state.last_order

    if not order:
        orders = load_orders()
        mine = orders[
            orders["Student_ID"].astype(str) == str(st.session_state.student_id)
        ]
        paid = mine[mine["Payment_Status"].astype(str).str.lower() == "paid"]
        if not paid.empty:
            order = paid.iloc[-1].to_dict()
            st.session_state.last_order = order

    if not order:
        st.info("No paid order/token available yet.")
        if st.button("🍔 Order Food"):
            reset_to("Food Menu")
        return

    if str(order.get("Payment_Status", "")).lower() != "paid":
        st.warning("Payment is required before token generation.")
        return

    st.markdown(
        f"""
        <div class="token-box">
            <div style="font-size:22px">🎟️ Your Token Number</div>
            <div style="font-size:64px;font-weight:800">{order["Token_No"]}</div>
            <div style="font-size:20px">Payment Successful ✅</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Order ID", order["Order_ID"])
    c2.metric("Student ID", order["Student_ID"])
    c3.metric("Queue Position", order["Queue_Position"])
    c4.metric("Wait Time", f"{order['Estimated_Wait_Minutes']} min")

    st.markdown(
        f"""
        **Student Name:** {order["Student_Name"]}  
        **Food Items:** {", ".join(
            f'{x.get("Food", "")} x{x.get("Quantity", 1)}'
            for x in parse_items(order.get("Food_Items", ""))
        )}  
        **Total Amount:** {money(order["Total_Amount"])}  
        **Order Date & Time:** {order["Order_Date"]}  
        **Pickup Counter:** {order["Pickup_Counter"]}  
        **Order Status:** {order["Order_Status"]}
        """
    )

    st.success(
        "This exact Token + Student ID + Order ID is also visible to the Canteen Owner."
    )

    if st.button("🧍 View Queue Position", use_container_width=True):
        reset_to("Queue")

    if st.button("📋 View Order Details", use_container_width=True):
        reset_to("Order Details")


# ============================================================
# QUEUE
# ============================================================

def queue_page():
    st.title("🧍 Queue Position")

    orders = load_orders()
    sid = st.session_state.student_id

    mine = orders[
        orders["Student_ID"].astype(str) == str(sid)
    ]

    if mine.empty:
        st.info("No orders found.")
        return

    active_statuses = ["Order Received", "Preparing", "Ready for Pickup"]
    active = mine[mine["Order_Status"].isin(active_statuses)]

    if active.empty:
        st.success("You currently have no active orders.")
        return

    current = active.iloc[-1]

    st.markdown(
        f"""
        ### 🎟️ Token: **{current["Token_No"]}**
        ### 📍 Queue Position: **{current["Queue_Position"]}**
        ### ⏱️ Estimated Waiting Time: **{current["Estimated_Wait_Minutes"]} minutes**
        ### 🍽️ Pickup Counter: **{current["Pickup_Counter"]}**
        ### 📦 Status: **{current["Order_Status"]}**
        """
    )

    st.info("The queue position is calculated when the paid order is created.")


# ============================================================
# OWNER DASHBOARD
# ============================================================

def owner_dashboard():
    st.title("👨‍🍳 Canteen Owner Dashboard")
    st.caption(
        "Every successfully paid student order appears here with the SAME token, "
        "Student ID and order details."
    )

    orders = load_orders()

    if orders.empty:
        st.info("No orders have been received yet.")
        return

    paid_orders = orders[
        orders["Payment_Status"].astype(str).str.lower() == "paid"
    ].copy()

    if paid_orders.empty:
        st.info("No paid orders have been received yet.")
        return

    # Latest orders first.
    paid_orders = paid_orders.iloc[::-1].reset_index(drop=True)

    # Summary cards
    active_statuses = ["Order Received", "Preparing", "Ready for Pickup"]
    active_count = len(
        paid_orders[paid_orders["Order_Status"].isin(active_statuses)]
    )
    completed_count = len(
        paid_orders[paid_orders["Order_Status"].isin(["Completed", "Picked Up"])]
    )
    revenue = pd.to_numeric(
        paid_orders["Total_Amount"], errors="coerce"
    ).fillna(0).sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🎟️ Paid Orders", len(paid_orders))
    c2.metric("⏳ Active Orders", active_count)
    c3.metric("✅ Completed", completed_count)
    c4.metric("💰 Revenue", money(revenue))

    st.divider()

    st.subheader("🎟️ Live Received Tokens")

    # Owner sees exactly the token generated for the student.
    display = paid_orders[
        [
            "Token_No", "Order_ID", "Student_ID", "Student_Name",
            "Food_Items", "Total_Amount", "Order_Date",
            "Payment_Status", "Order_Status", "Queue_Position",
            "Estimated_Wait_Minutes", "Pickup_Counter"
        ]
    ].copy()

    display["Food_Items"] = display["Food_Items"].apply(
        lambda x: ", ".join(
            f"{item.get('Food', '')} x{item.get('Quantity', 1)}"
            for item in parse_items(x)
        )
    )
    display["Total_Amount"] = display["Total_Amount"].apply(money)
    display["Estimated_Wait_Minutes"] = (
        display["Estimated_Wait_Minutes"].astype(str) + " min"
    )

    display = display.rename(columns={
        "Token_No": "TOKEN",
        "Order_ID": "ORDER ID",
        "Student_ID": "STUDENT ID",
        "Student_Name": "STUDENT NAME",
        "Food_Items": "FOOD ITEMS",
        "Total_Amount": "AMOUNT",
        "Order_Date": "ORDER DATE & TIME",
        "Payment_Status": "PAYMENT",
        "Order_Status": "ORDER STATUS",
        "Queue_Position": "QUEUE",
        "Estimated_Wait_Minutes": "WAIT",
        "Pickup_Counter": "PICKUP"
    })

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True
    )

    st.divider()

    # Detailed view of the latest received order
    st.subheader("📋 Latest Student Order")

    latest = paid_orders.iloc[0]

    a, b, c, d = st.columns(4)
    a.metric("🎟️ Token", latest["Token_No"])
    b.metric("👤 Student ID", latest["Student_ID"])
    c.metric("🧾 Order ID", latest["Order_ID"])
    d.metric("💰 Amount", money(latest["Total_Amount"]))

    st.markdown(
        f"""
        **Student Name:** {latest["Student_Name"]}  
        **Order Date & Time:** {latest["Order_Date"]}  
        **Payment:** {latest["Payment_Status"]}  
        **Queue Position:** {latest["Queue_Position"]}  
        **Estimated Wait:** {latest["Estimated_Wait_Minutes"]} minutes  
        **Pickup Counter:** {latest["Pickup_Counter"]}  
        **Order Status:** {latest["Order_Status"]}
        """
    )

    st.markdown("### 🍔 Food Items")
    for item in parse_items(latest["Food_Items"]):
        st.write(
            f'{item.get("Emoji", "🍽️")} **{item.get("Food", "")}** '
            f'x {item.get("Quantity", 1)} — '
            f'{money(float(item.get("Price", 0)) * int(item.get("Quantity", 1)))}'
        )

    st.markdown("### 🔄 Update Order Status")

    status_options = [
        "Order Received",
        "Preparing",
        "Ready for Pickup",
        "Completed",
        "Picked Up"
    ]

    current_status = latest["Order_Status"]
    status_index = (
        status_options.index(current_status)
        if current_status in status_options else 0
    )

    new_status = st.selectbox(
        "Order Status",
        status_options,
        index=status_index,
        key=f"owner_status_{latest['Order_ID']}"
    )

    if st.button(
        "💾 Update Latest Order Status",
        type="primary",
        use_container_width=True
    ):
        all_orders = load_orders()
        mask = all_orders["Order_ID"].astype(str) == str(latest["Order_ID"])
        all_orders.loc[mask, "Order_Status"] = new_status
        save_orders(all_orders)
        st.success(
            f'Token {latest["Token_No"]} updated to "{new_status}".'
        )
        st.rerun()


# ============================================================
# RECEIVED ORDERS
# ============================================================

def received_orders_page():
    st.title("📦 Received Orders")
    st.caption(
        "The owner receives the same token generated on the student side, "
        "along with Student ID and complete order details."
    )

    orders = load_orders()

    if orders.empty:
        st.info("No orders received yet.")
        return

    paid = orders[
        orders["Payment_Status"].astype(str).str.lower() == "paid"
    ].copy()

    if paid.empty:
        st.info("No paid orders received yet.")
        return

    paid = paid.iloc[::-1].reset_index(drop=True)

    for _, row in paid.iterrows():
        with st.container(border=True):
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("🎟️ TOKEN", row["Token_No"])
            c2.metric("👤 STUDENT ID", row["Student_ID"])
            c3.metric("🧾 ORDER ID", row["Order_ID"])
            c4.metric("💰 AMOUNT", money(row["Total_Amount"]))
            c5.metric("📍 QUEUE", row["Queue_Position"])

            st.write(f"**Student Name:** {row['Student_Name']}")
            st.write(f"**Order Date & Time:** {row['Order_Date']}")
            st.write(f"**Payment:** {row['Payment_Status']}")
            st.write(f"**Status:** {row['Order_Status']}")
            st.write(f"**Pickup:** {row['Pickup_Counter']}")

            items = parse_items(row["Food_Items"])
            item_text = ", ".join(
                f"{x.get('Food', '')} x{x.get('Quantity', 1)}"
                for x in items
            )
            st.write(f"**Food Items:** {item_text}")

    if st.button("🔄 Refresh Orders", use_container_width=True):
        st.rerun()


# ============================================================
# ORDER DETAILS
# ============================================================

def order_details_page():
    st.title("📋 Order Details")

    orders = load_orders()
    sid = st.session_state.student_id

    mine = orders[
        orders["Student_ID"].astype(str) == str(sid)
    ].copy()

    if mine.empty:
        st.info("No orders yet.")
        return

    selected_id = st.selectbox(
        "Select Order",
        mine["Order_ID"].astype(str).tolist(),
        index=len(mine) - 1
    )

    row = mine[mine["Order_ID"].astype(str) == str(selected_id)].iloc[0]

    items = parse_items(row["Food_Items"])

    item_text = []
    for item in items:
        item_text.append(
            f'{item.get("Emoji", "🍽️")} {item.get("Food", "")} '
            f'x {item.get("Quantity", 1)}'
        )

    st.success(
        f"Payment Status: {row['Payment_Status']} | "
        f"Order Status: {row['Order_Status']}"
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Serial No.", str(list(mine["Order_ID"]).index(selected_id) + 1).zfill(3))
    c2.metric("Order ID", row["Order_ID"])
    c3.metric("Token No.", row["Token_No"])

    c1, c2, c3 = st.columns(3)
    c1.metric("Queue Position", row["Queue_Position"])
    c2.metric("Wait Time", f"{row['Estimated_Wait_Minutes']} minutes")
    c3.metric("Total Amount", money(row["Total_Amount"]))

    st.write("")
    st.markdown("### 🍔 Food Items")
    for text_item in item_text:
        st.write(text_item)

    st.write("")
    st.markdown("### 👤 Student & Pickup Information")
    st.write(f"**Student ID:** {row['Student_ID']}")
    st.write(f"**Student Name:** {row['Student_Name']}")
    st.write(f"**Order Date & Time:** {row['Order_Date']}")
    st.write(f"**Pickup Counter:** {row['Pickup_Counter']}")

    if st.session_state.last_payment:
        payment = st.session_state.last_payment
        if payment.get("Order_ID") == row["Order_ID"]:
            st.write(f"**Payment Method:** {payment.get('Payment_Method', '-')}")
            st.write(f"**Transaction ID:** {payment.get('Transaction_ID', '-')}")


# ============================================================
# ANNUAL REPORT DATA
# ============================================================

def build_report_data(year_label):
    """
    Converts an academic year such as 2026-2027 into a date range:
    2026-04-01 through 2027-03-31.
    """
    start_year = int(year_label.split("-")[0])
    start_date = pd.Timestamp(year=start_year, month=4, day=1)
    end_date = pd.Timestamp(year=start_year + 1, month=3, day=31, hour=23, minute=59, second=59)

    orders = load_orders()

    if orders.empty:
        filtered = orders.copy()
    else:
        filtered = orders.copy()
        filtered["Order_Date_DT"] = pd.to_datetime(
            filtered["Order_Date"], errors="coerce"
        )
        filtered = filtered[
            (filtered["Order_Date_DT"] >= start_date)
            & (filtered["Order_Date_DT"] <= end_date)
        ].copy()

    return filtered, start_date, end_date


def food_sales_table(filtered_orders):
    records = []

    for _, row in filtered_orders.iterrows():
        for item in parse_items(row.get("Food_Items", "")):
            quantity = int(item.get("Quantity", 0) or 0)
            price = float(item.get("Price", 0) or 0)

            records.append({
                "Food Item": f"{item.get('Emoji', '')} {item.get('Food', '')}".strip(),
                "Quantity Sold": quantity,
                "Revenue": quantity * price
            })

    if not records:
        return pd.DataFrame(
            columns=["Food Item", "Quantity Sold", "Revenue"]
        )

    df = pd.DataFrame(records)
    df = df.groupby("Food Item", as_index=False).agg({
        "Quantity Sold": "sum",
        "Revenue": "sum"
    })
    df["Revenue"] = df["Revenue"].round(2)
    return df.sort_values("Revenue", ascending=False)


def monthly_sales_table(filtered_orders, start_year):
    months = [
        ("Apr", 4, start_year),
        ("May", 5, start_year),
        ("Jun", 6, start_year),
        ("Jul", 7, start_year),
        ("Aug", 8, start_year),
        ("Sep", 9, start_year),
        ("Oct", 10, start_year),
        ("Nov", 11, start_year),
        ("Dec", 12, start_year),
        ("Jan", 1, start_year + 1),
        ("Feb", 2, start_year + 1),
        ("Mar", 3, start_year + 1),
    ]

    rows = []

    for name, month, year in months:
        if filtered_orders.empty:
            month_orders = filtered_orders
        else:
            month_orders = filtered_orders[
                (filtered_orders["Order_Date_DT"].dt.year == year)
                & (filtered_orders["Order_Date_DT"].dt.month == month)
            ]

        revenue = pd.to_numeric(
            month_orders["Total_Amount"], errors="coerce"
        ).fillna(0).sum()

        rows.append({
            "Month": name,
            "Orders": len(month_orders),
            "Revenue": revenue
        })

    return pd.DataFrame(rows)


# ============================================================
# ANNUAL REPORT EXPORTS
# ============================================================

def create_excel_report(year_label, filtered_orders, food_df, monthly_df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary = pd.DataFrame([
            ["Academic Year", year_label],
            ["Total Orders", len(filtered_orders)],
            [
                "Completed Orders",
                len(filtered_orders[
                    filtered_orders["Order_Status"].isin(["Completed", "Picked Up"])
                ])
            ],
            [
                "Active Orders",
                len(filtered_orders[
                    filtered_orders["Order_Status"].isin(
                        ["Order Received", "Preparing", "Ready for Pickup"]
                    )
                ])
            ],
            [
                "Total Revenue",
                pd.to_numeric(
                    filtered_orders["Total_Amount"], errors="coerce"
                ).fillna(0).sum()
            ],
            [
                "Food Items Sold",
                int(food_df["Quantity Sold"].sum()) if not food_df.empty else 0
            ],
            [
                "Students Served",
                filtered_orders["Student_ID"].nunique()
                if not filtered_orders.empty else 0
            ],
        ], columns=["Metric", "Value"])

        summary.to_excel(writer, sheet_name="Summary", index=False)
        food_df.to_excel(writer, sheet_name="Food Sales", index=False)
        monthly_df.to_excel(writer, sheet_name="Monthly Report", index=False)

        if not filtered_orders.empty:
            display_orders = filtered_orders.copy()
            if "Order_Date_DT" in display_orders.columns:
                display_orders = display_orders.drop(columns=["Order_Date_DT"])
            display_orders.to_excel(writer, sheet_name="Orders", index=False)

        # Style workbook
        workbook = writer.book
        for ws in workbook.worksheets:
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")

            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except Exception:
                        pass
                ws.column_dimensions[column_letter].width = min(max_length + 3, 35)

    output.seek(0)
    return output.getvalue()


def create_pdf_report(year_label, filtered_orders, food_df, monthly_df):
    if not REPORTLAB_OK:
        return None

    output = BytesIO()

    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleCustom",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=22,
        leading=26
    )

    story = []

    story.append(Paragraph(
        "SMART CANTEEN — ANNUAL REPORT",
        title_style
    ))
    story.append(Paragraph(
        f"Academic Year: {year_label}",
        styles["Heading2"]
    ))
    story.append(Spacer(1, 12))

    total_revenue = pd.to_numeric(
        filtered_orders["Total_Amount"], errors="coerce"
    ).fillna(0).sum()

    completed = len(
        filtered_orders[
            filtered_orders["Order_Status"].isin(["Completed", "Picked Up"])
        ]
    )

    active = len(
        filtered_orders[
            filtered_orders["Order_Status"].isin(
                ["Order Received", "Preparing", "Ready for Pickup"]
            )
        ]
    )

    food_count = int(food_df["Quantity Sold"].sum()) if not food_df.empty else 0
    students = filtered_orders["Student_ID"].nunique() if not filtered_orders.empty else 0

    summary_data = [
        ["Metric", "Value"],
        ["Students Served", str(students)],
        ["Total Orders", str(len(filtered_orders))],
        ["Completed Orders", str(completed)],
        ["Active Orders", str(active)],
        ["Food Items Sold", str(food_count)],
        ["Total Revenue", money(total_revenue)],
    ]

    table = Table(summary_data, colWidths=[220, 140])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e2e8f0")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 1), (1, -1), "CENTER"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
    ]))

    story.append(Paragraph("Key Statistics", styles["Heading2"]))
    story.append(table)
    story.append(Spacer(1, 15))

    # Food sales
    story.append(Paragraph("Food-wise Sales", styles["Heading2"]))

    food_data = [["Food Item", "Quantity Sold", "Revenue"]]
    if food_df.empty:
        food_data.append(["No sales data", "0", "₹0"])
    else:
        for _, row in food_df.iterrows():
            food_data.append([
                str(row["Food Item"]),
                str(int(row["Quantity Sold"])),
                money(row["Revenue"])
            ])

    food_table = Table(food_data, colWidths=[260, 130, 130])
    food_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e2e8f0")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
    ]))

    story.append(food_table)
    story.append(Spacer(1, 15))

    # Monthly report
    story.append(Paragraph("Monthly Report", styles["Heading2"]))

    monthly_data = [["Month", "Orders", "Revenue"]]
    for _, row in monthly_df.iterrows():
        monthly_data.append([
            str(row["Month"]),
            str(int(row["Orders"])),
            money(row["Revenue"])
        ])

    monthly_table = Table(monthly_data, colWidths=[100, 100, 130])
    monthly_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e2e8f0")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
    ]))

    story.append(monthly_table)
    story.append(Spacer(1, 15))

    story.append(Paragraph(
        f"Report generated on {datetime.now().strftime('%d-%m-%Y %I:%M %p')}",
        styles["Normal"]
    ))

    doc.build(story)

    output.seek(0)
    return output.getvalue()


# ============================================================
# ANNUAL REPORT PAGE
# ============================================================

def annual_report_page():
    st.title("📊 Annual Report")
    st.caption(
        "The report is calculated from the actual paid orders stored by this app."
    )

    current_year = datetime.now().year

    years = [
        f"{current_year}-{current_year + 1}",
        f"{current_year - 1}-{current_year}",
        f"{current_year - 2}-{current_year - 1}",
        f"{current_year - 3}-{current_year - 2}",
    ]

    year_label = st.selectbox("📅 Select Academic Year", years)

    filtered, start_date, end_date = build_report_data(year_label)

    # Revenue and counts
    if filtered.empty:
        total_revenue = 0
        students = 0
        food_df = food_sales_table(filtered)
    else:
        total_revenue = pd.to_numeric(
            filtered["Total_Amount"], errors="coerce"
        ).fillna(0).sum()

        students = filtered["Student_ID"].nunique()
        food_df = food_sales_table(filtered)

    completed = len(
        filtered[
            filtered["Order_Status"].isin(["Completed", "Picked Up"])
        ]
    )

    active = len(
        filtered[
            filtered["Order_Status"].isin(
                ["Order Received", "Preparing", "Ready for Pickup"]
            )
        ]
    )

    food_items_sold = (
        int(food_df["Quantity Sold"].sum())
        if not food_df.empty else 0
    )

    st.info(
        f"Academic year period: {start_date.strftime('%d-%m-%Y')} "
        f"to {end_date.strftime('%d-%m-%Y')}"
    )

    # KPI cards
    c1, c2, c3 = st.columns(3)
    c1.metric("👨‍🎓 Total Students", students)
    c2.metric("🛒 Total Orders", len(filtered))
    c3.metric("🍔 Food Items Sold", food_items_sold)

    c1, c2, c3 = st.columns(3)
    c1.metric("✅ Completed Orders", completed)
    c2.metric("⏳ Active Orders", active)
    c3.metric("💰 Total Revenue", money(total_revenue))

    st.divider()

    # Food analysis
    st.subheader("🍔 Food Analysis")

    if food_df.empty:
        st.info("No paid order data is available for this academic year.")
    else:
        display_food = food_df.copy()
        display_food["Revenue"] = display_food["Revenue"].apply(money)
        st.dataframe(
            display_food,
            use_container_width=True,
            hide_index=True
        )

        st.subheader("📊 Food-wise Quantity Sold")
        chart_food = food_df.set_index("Food Item")[["Quantity Sold"]]
        st.bar_chart(chart_food)

        st.subheader("💰 Food-wise Revenue")
        chart_revenue = food_df.set_index("Food Item")[["Revenue"]]
        st.bar_chart(chart_revenue)

    # Monthly report
    st.subheader("📅 Monthly Report")

    monthly_df = monthly_sales_table(
        filtered,
        int(year_label.split("-")[0])
    )

    monthly_display = monthly_df.copy()
    monthly_display["Revenue"] = monthly_display["Revenue"].apply(money)

    st.dataframe(
        monthly_display,
        use_container_width=True,
        hide_index=True
    )

    st.subheader("📈 Monthly Revenue")
    st.line_chart(
        monthly_df.set_index("Month")[["Revenue"]]
    )

    st.subheader("📊 Monthly Orders")
    st.bar_chart(
        monthly_df.set_index("Month")[["Orders"]]
    )

    # Status chart
    st.subheader("🟢 Completed vs Active Orders")

    status_df = pd.DataFrame({
        "Status": ["Completed", "Active"],
        "Orders": [completed, active]
    })

    st.bar_chart(status_df.set_index("Status"))

    # Download reports
    st.divider()
    st.subheader("📥 Download Annual Report")

    excel_bytes = create_excel_report(
        year_label,
        filtered,
        food_df,
        monthly_df
    )

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "📥 Download Excel Report",
            data=excel_bytes,
            file_name=f"Smart_Canteen_Annual_Report_{year_label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        if REPORTLAB_OK:
            pdf_bytes = create_pdf_report(
                year_label,
                filtered,
                food_df,
                monthly_df
            )

            st.download_button(
                "📄 Download PDF Report",
                data=pdf_bytes,
                file_name=f"Smart_Canteen_Annual_Report_{year_label}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        else:
            st.warning(
                "PDF download needs reportlab. Install it with: "
                "py -m pip install reportlab"
            )


# ============================================================
# MAIN APP
# ============================================================

if not st.session_state.logged_in:
    login_page()
else:
    sidebar()

    page = st.session_state.page

    if st.session_state.user_role == "owner":
        if page == "Owner Dashboard":
            owner_dashboard()
        elif page == "Received Orders":
            received_orders_page()
        elif page == "Annual Report":
            annual_report_page()
        else:
            owner_dashboard()
    else:
        if page == "Dashboard":
            dashboard()
        elif page == "Food Menu":
            food_menu()
        elif page == "Cart":
            cart_page()
        elif page == "Checkout":
            checkout_page()
        elif page == "Payment":
            payment_page()
        elif page == "Token":
            token_page()
        elif page == "Queue":
            queue_page()
        elif page == "Order Details":
            order_details_page()
        elif page == "Annual Report":
            annual_report_page()
        else:
            dashboard()
