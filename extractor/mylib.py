import re


COUNTRIES_WITH_STATES = {
    "United States": [
        "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
        "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho",
        "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana",
        "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota",
        "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada",
        "New Hampshire", "New Jersey", "New Mexico", "New York",
        "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon",
        "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota",
        "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington",
        "West Virginia", "Wisconsin", "Wyoming",
    ],

    "Germany": [
        "Baden-Württemberg", "Bavaria", "Berlin", "Brandenburg", "Bremen",
        "Hamburg", "Hesse", "Lower Saxony", "Mecklenburg-Vorpommern",
        "North Rhine-Westphalia", "Rhineland-Palatinate", "Saarland",
        "Saxony", "Saxony-Anhalt", "Schleswig-Holstein", "Thuringia",
    ],

    "India": [
        "Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar",
        "Chhattisgarh", "Goa", "Gujarat", "Haryana", "Himachal Pradesh",
        "Jharkhand", "Karnataka", "Kerala", "Madhya Pradesh", "Maharashtra",
        "Manipur", "Meghalaya", "Mizoram", "Nagaland", "Odisha", "Punjab",
        "Rajasthan", "Sikkim", "Tamil Nadu", "Telangana", "Tripura",
        "Uttar Pradesh", "Uttarakhand", "West Bengal",
        "Andaman and Nicobar Islands", "Chandigarh",
        "Dadra and Nagar Haveli and Daman and Diu", "Delhi",
        "Jammu and Kashmir", "Ladakh", "Lakshadweep", "Puducherry",
    ],

    "United Kingdom": [
        "England", "Scotland", "Wales", "Northern Ireland",
    ],

    "Canada": [
        "Alberta", "British Columbia", "Manitoba", "New Brunswick",
        "Newfoundland and Labrador", "Nova Scotia", "Ontario",
        "Prince Edward Island", "Quebec", "Saskatchewan",
        "Northwest Territories", "Nunavut", "Yukon",
    ],
}


COUNTRY_NAME_ALIASES = {
    "united states of america": "United States",
    "united states": "United States",
    "usa": "United States",
    "u.s.a.": "United States",
    "u.s.a": "United States",
    "u.s.": "United States",
    "u.s": "United States",
    "us": "United States",

    "germany": "Germany",

    "india": "India",

    "united kingdom": "United Kingdom",
    "great britain": "United Kingdom",
    "u.k.": "United Kingdom",
    "u.k": "United Kingdom",
    "uk": "United Kingdom",

    "canada": "Canada",
}


def normalize_location_text(value):
    """
    Normalize location/state/country text before matching.
    """
    if not value:
        return ""

    value = str(value).replace("\xa0", " ").strip()
    value = re.sub(r"\s+", " ", value)

    return value


def find_explicit_country(location):
    """
    Find a country directly written in a location string.

    Examples:
        Mumbai, Maharashtra, India -> India
        Austin, Texas, USA -> United States
        London, UK -> United Kingdom
    """
    location = normalize_location_text(location)

    if not location:
        return ""

    aliases = sorted(
        COUNTRY_NAME_ALIASES.items(),
        key=lambda item: len(item[0]),
        reverse=True,
    )

    for alias, country in aliases:
        pattern = (
            r"(?<![A-Za-z])"
            + re.escape(alias)
            + r"(?![A-Za-z])"
        )

        if re.search(pattern, location, flags=re.IGNORECASE):
            return country

    return ""


def find_country_from_state(state_or_location):
    """
    Find country by matching a state/province/territory.

    Examples:
        Maharashtra -> India
        Austin, Texas -> United States
        Toronto, Ontario -> Canada
        Bavaria -> Germany
        England -> United Kingdom
    """
    text = normalize_location_text(state_or_location)

    if not text:
        return ""

    matched_countries = set()

    for country, states in COUNTRIES_WITH_STATES.items():
        for state in sorted(states, key=len, reverse=True):
            pattern = (
                r"(?<![A-Za-z])"
                + re.escape(state)
                + r"(?![A-Za-z])"
            )

            if re.search(pattern, text, flags=re.IGNORECASE):
                matched_countries.add(country)
                break

    if len(matched_countries) == 1:
        return next(iter(matched_countries))

    return ""


def infer_country_from_location(
    location="",
    state="",
    existing_country="",
):
    """
    Find country using this priority:

    1. Keep existing country if already extracted.
    2. Find country directly inside location.
    3. Find country directly inside state.
    4. Match state/province/territory with COUNTRIES_WITH_STATES.
    5. Return blank if country cannot be found reliably.

    This function is common for all job portals.
    """

    existing_country = normalize_location_text(existing_country)

    if existing_country:
        return existing_country

    location = normalize_location_text(location)
    state = normalize_location_text(state)

    country = find_explicit_country(location)
    if country:
        return country

    country = find_explicit_country(state)
    if country:
        return country

    search_text = " ".join(
        value
        for value in [location, state]
        if value
    )

    country = find_country_from_state(search_text)
    if country:
        return country

    return ""
