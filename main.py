import subprocess
import time
import sys

def run_demo():
    print("==================================================")
    print("🚀 INITIATING EMAIL AUTOMATION CAMPAIGN PIPELINE")
    print("==================================================\n")

    # Step 1: Execute the Bulk Mailer
    print("[SYSTEM] Reading ceo_data.xlsx and dispatching cold emails...")
    try:
        # This runs your bulk_mailer.py script automatically
        subprocess.run([sys.executable, "bulk_mailer.py"], check=True)
        print("\n✅ [SUCCESS] All outreach emails dispatched successfully.")
    except subprocess.CalledProcessError:
        print("\n❌ [ERROR] The bulk mailer encountered an issue.")
        return

    print("-" * 50)
    print("[SYSTEM] Transitioning to Phase 2: Inbox Listener")
    print("-" * 50)
    time.sleep(2)

    # Step 2: Execute the Auto-Reply Listener
    print("[SYSTEM] Starting IMAP Listener. Waiting for Calendly bookings...\n")
    try:
        # This runs your auto_reply.py script automatically
        subprocess.run([sys.executable, "auto_reply.py"], check=True)
    except KeyboardInterrupt:
        print("\n🛑 [SYSTEM] Listener stopped by user.")

if __name__ == "__main__":
    run_demo()