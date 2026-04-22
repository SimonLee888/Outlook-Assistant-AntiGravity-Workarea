from datetime import datetime, timedelta


def get_vehicle_data(start_ms, end_ms):
    return [
        {
            "id": "vehicle-1",
            "name": "Sanitized Vehicle",
            "locations": [
                {"latitude": 0.0, "longitude": 0.0, "speedMilesPerHour": 0, "timeMs": start_ms}
            ],
        }
    ]


def save_to_postgres(data):
    return len(data)


def fetch_and_store_historical_data(start_date, end_date):
    current_start = start_date
    while current_start < end_date:
        current_end = min(current_start + timedelta(minutes=30), end_date)
        start_ms = int(current_start.timestamp() * 1000)
        end_ms = int(current_end.timestamp() * 1000)

        vehicle_data = get_vehicle_data(start_ms, end_ms)
        if vehicle_data:
            save_to_postgres(vehicle_data)

        current_start = current_end


if __name__ == "__main__":
    fetch_and_store_historical_data(datetime(2024, 1, 1), datetime.utcnow())
