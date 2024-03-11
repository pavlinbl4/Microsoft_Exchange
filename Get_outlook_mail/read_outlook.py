import exchangelib
from bs4 import BeautifulSoup
from dotenv import load_dotenv
import os


def make_soup(letter):
    soup = BeautifulSoup(letter, 'lxml')
    return soup.find('body').text.strip()


def authorization():
    load_dotenv()

    credentials = exchangelib.Credentials(
        username=os.environ.get('username'),
        password=os.environ.get('password')
    )

    config = exchangelib.Configuration(
        server=os.environ.get('server'),
        credentials=credentials
    )

    account = exchangelib.Account(
        primary_smtp_address=os.environ.get('primary_smtp_address'),
        credentials=credentials,
        autodiscover=False,
        config=config
    )

    return account
