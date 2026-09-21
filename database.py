import pandas as pd
import os

DATA_FOLDER = "data"

USERS_FILE = os.path.join(DATA_FOLDER, "users.csv")
MENU_FILE = os.path.join(DATA_FOLDER, "menu.csv")
ORDERS_FILE = os.path.join(DATA_FOLDER, "orders.csv")


def load_users():
    return pd.read_csv(USERS_FILE)


def save_users(df):
    df.to_csv(USERS_FILE, index=False)


def load_menu():
    return pd.read_csv(MENU_FILE)


def save_menu(df):
    df.to_csv(MENU_FILE, index=False)


def load_orders():
    return pd.read_csv(ORDERS_FILE)


def save_orders(df):
    df.to_csv(ORDERS_FILE, index=False)