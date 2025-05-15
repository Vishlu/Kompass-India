import re
import smtplib
import dns.resolver
from time import sleep
from collections import deque

def verify_emails(emails):
    def validate_email_syntax(email):
        email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        return bool(re.match(email_regex, email))

    def get_mx_records(domain):
        try:
            result = dns.resolver.query(domain, 'MX')
            return str(result[0].exchange)
        except dns.resolver.NoAnswer:
            return None

    def check_email_active(email):
        if not validate_email_syntax(email):
            return False, "Invalid email syntax."

        domain = email.split('@')[-1]

        if not domain:
            return False, "Domain is empty."

        if len(domain) > 253:
            return False, "Domain name is too long."

        labels = domain.split('.')
        for label in labels:
            if len(label) > 63:
                return False, f"Label '{label}' is too long."

        sleep(1)

        mail_server = get_mx_records(domain)
        if not mail_server:
            return False, f"Domain '{domain}' does not have MX records."

        try:
            server = smtplib.SMTP(mail_server, timeout=10)
            server.set_debuglevel(0)
            server.helo()
            code, message = server.mail("test@example.com")
            if code == 250:
                code, message = server.rcpt(email)
                if code == 250:
                    server.quit()
                    return True, f"Email '{email}' is valid and active."
                else:
                    server.quit()
                    return False, f"Email '{email}' is not valid (RCPT TO failed)."
            else:
                server.quit()
                return False, f"Error checking email: {message}"
        except smtplib.SMTPException as e:
            return False, f"Error checking email: {e}"

    results = []
    queue = deque(emails)

    while queue:
        email = queue.popleft()
        try:
            is_active, message = check_email_active(email)
            results.append((email, is_active, message))
        except Exception as e:
            results.append((email, False, f"Error: {e}"))
            print(f"Error verifying email: {email} ({e})")

        sleep(2)

    return results

if __name__ == "__main__":
    SingleQuickEmails_Verify = [
        "akshfggvay@kompassindia.com",
    ]

    results = verify_emails(SingleQuickEmails_Verify)
    for email, is_active, message in results:
        print(f"Email: {email}, Status: {'Active' if is_active else 'Inactive'}, Message: {message}")