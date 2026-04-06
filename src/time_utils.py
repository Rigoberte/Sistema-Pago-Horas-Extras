import pandas as pd


def round_timestamp_to_nearest_half_hour(ts: pd.Timestamp) -> pd.Timestamp:
    total_minutes = ts.hour * 60 + ts.minute + ts.second / 60
    rounded_minutes = int((total_minutes + 15) // 30) * 30
    rounded_minutes %= 24 * 60

    hour = rounded_minutes // 60
    minute = rounded_minutes % 60
    return ts.replace(hour=hour, minute=minute, second=0, microsecond=0)
