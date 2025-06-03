
import requests
from dotenv import load_dotenv
import os
from pprint import pprint

load_dotenv()

def get_current_weather():
    city = input("\nEnter the city name:\n")
    request_url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={os.getenv('API_KEY')}&units=metric"

    response = requests.get(request_url).json()

    # pprint(response)
    print("\n")
    print(f"City: {city}")
    print(f"Temperature: {response['main']['temp']}°C")
    print(f"Weather: {response['weather'][0]['description']}")
    print(f"Humidity: {response['main']['humidity']}%")
    print(f"Wind Speed: {response['wind']['speed']} m/s")           


if __name__ == "__main__":
    get_current_weather()


