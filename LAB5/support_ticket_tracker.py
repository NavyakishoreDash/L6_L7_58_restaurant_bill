"""Support Ticket Tracker module.

Provides functionality to create, assign, update, search, and persist
customer support tickets with priority and category tracking.
"""

import json
import logging
import os

TICKETS = {}


def create_ticket(
    title, description, reporter, priority="medium", **kwargs
):
    """Create a new support ticket and record an entry in the audit log."""
    if not title:
        return None
    category = kwargs.get("category", "general")
    audit_log = kwargs.get("audit_log")
    if audit_log is None:
        audit_log = []

    ticket_id = str(len(TICKETS) + 1)
    ticket = {
        "id": ticket_id,
        "title": title,
        "description": description,
        "priority": priority,
        "category": category,
        "status": "open",
        "reporter": reporter,
        "assignee": None,
    }
    TICKETS[ticket_id] = ticket
    audit_log.append(f"Created ticket {ticket_id} for {reporter}")
    return ticket


def assign_ticket(ticket_id, agent, notify=False):
    """Assign a support ticket to a specific agent with optional alert."""
    if ticket_id not in TICKETS:
        logging.warning("Ticket ID %s not found for assignment", ticket_id)
        return False
    TICKETS[ticket_id]["assignee"] = agent
    if notify:
        logging.info("Agent %s notified for ticket %s", agent, ticket_id)
    return True


def update_status(ticket_id, status):
    """Update the resolution status of an existing support ticket."""
    if ticket_id in TICKETS:
        TICKETS[ticket_id]["status"] = status
        return TICKETS[ticket_id]
    return None


def search_tickets(status=None, priority=None):
    """Filter and return tickets matching the given status and priority."""
    results = []
    for ticket in TICKETS.values():
        if status is not None and ticket["status"] != status:
            continue
        if priority and ticket["priority"] != priority:
            continue
        results.append(ticket)
    return results


def save_tickets(path="tickets.json"):
    """Persist the current tickets database to a JSON file safely."""
    with open(path, "w", encoding="utf-8") as file:
        file.write(json.dumps(TICKETS, indent=2))


def load_tickets(path="tickets.json"):
    """Load tickets from a JSON file into memory without global keyword."""
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as file:
        loaded = json.loads(file.read())
        TICKETS.clear()
        TICKETS.update(loaded)
    return TICKETS


def print_ticket_report():
    """Display a formatted report of all recorded support tickets."""
    print("Support Tickets")
    if not TICKETS:
        print("No tickets yet")
    for ticket_id, ticket in TICKETS.items():
        print(ticket_id, ticket["title"], ticket["status"])


def close_ticket(ticket_id):
    """Close an existing support ticket by ID in snake_case style."""
    if ticket_id in TICKETS:
        TICKETS[ticket_id]["status"] = "closed"


def is_open(ticket_id):
    """Check whether a support ticket is currently open or in progress."""
    ticket = TICKETS.get(ticket_id)
    if not ticket:
        return False
    return ticket["status"] in ("open", "in_progress")


def main():
    """Execute the support ticket tracker demonstration workflow."""
    logging.basicConfig(level=logging.INFO)
    create_ticket(
        "Cannot sign in",
        "Password reset link fails",
        "alex@example.com",
    )
    create_ticket(
        "Billing question",
        "Charged twice",
        "lee@example.com",
        "high",
    )
    assign_ticket("1", "agent-01")
    update_status("1", "in_progress")
    print_ticket_report()
    logging.info("Tickets created: %d", len(TICKETS))
    print("Ticket 1 open?", is_open("1"))
    ticket_id = input("Ticket id to close (Enter to skip): ")
    if ticket_id:
        close_ticket(ticket_id)
    save_tickets()
    if os.environ.get("RUN_UNSAFE_DEMO") == "1":
        print("unsafe demo")


if __name__ == "__main__":
    main()
