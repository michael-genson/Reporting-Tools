def get_map():
    return {
        "1st attempt": {"follow-up": 2, "weight": 1.75},
        "2nd attempt": {"follow-up": 2, "weight": 1.75},
        "3rd attempt": {"follow-up": 2, "weight": 1.75},
        "escalated": {"follow-up": 1, "weight": 2},
        "new": {"follow-up": 1, "weight": 2},
        "new email received": {"follow-up": 1, "weight": 2},
        "open": {"follow-up": 1, "weight": 2},
        "re-opened": {"follow-up": 1, "weight": 2},
        "waiting on customer": {"follow-up": 2, "weight": 1.75},
        "waiting on development": {"follow-up": 14, "weight": 1},
        "waiting on ols": {"follow-up": 2, "weight": 1.75},
        "waiting on other follett department": {"follow-up": 2, "weight": 1.75},
        "waiting on 3rd party": {"follow-up": 2, "weight": 1.75},
        "working": {"follow-up": 7, "weight": 1.25},
    }
