import re


def get_valid_email():
    while True:
        recipient_email = input("Enter the recipient's email address: ").strip()
        if re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', recipient_email):
            return recipient_email
        else:
            print("Invalid email address. Please enter a valid email address.")
