import requests

import keys


def load_filters():
    filters = {}

    headers = {
        "QB-Realm-Hostname": keys.QB_REALM_HOSTNAME,
        "User-Agent": keys.USER_AGENT,
        "Authorization": keys.TOKEN,
    }

    qb_ids = [
        [
            keys.DOMAINS_TABLE_ID,
            keys.DOMAINS_COLUMN_ID,
            "domains",
        ],
        [
            keys.EMAILS_TABLE_ID,
            keys.EMAILS_COLUMN_ID,
            "emails",
        ],
        [
            keys.KEYWORDS_TABLE_ID,
            keys.KEYWORDS_COLUMN_ID,
            "keywords",
        ],
    ]

    for table_id, column_id, name in qb_ids:
        body = {
            "from": table_id,
            "select": [column_id],
            "where": f"{{{column_id}.XEX.''}}",
            "options": {"skip": 0, "top": 0, "compareWithAppLocalTime": False},
        }

        # Make API call
        try:
            r = requests.post(
                f"https://api.quickbase.com/v1/records/query",
                headers=headers,
                json=body,
            )
        except requests.ConnectionError:
            print("Could not make connection to Quickbase. Check internet connection.")
            return None

        # Check if request successfully returned
        try:
            r.raise_for_status()
        except requests.HTTPError:
            print("Quickbase servers returned invalid HTTP code.")
            print(f"{r.ok=}")
            print(f"{r.status_code=}")
            return None

        # Access the data we want, add to filters dict
        try:
            filters[name] = set(
                record[column_id]["value"] for record in r.json()["data"]
            )
        except KeyError:
            print("JSON returned from Quickbase API could not be decoded.")
            return None

    return filters


if __name__ == "__main__":
    print("This module exists to load data from the Quickbase API. Import to use.")
