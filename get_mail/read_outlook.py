import exchangelib
from bs4 import BeautifulSoup
from icecream import ic
from colorama import Fore, Back, Style
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


def read_email():
    inbox = authorization().inbox
    for item in inbox.all().order_by('-datetime_received')[:5]:
        print("*" * 100)
        print()
        print(Fore.RED + item.sender.name)
        print(Fore.GREEN + item.sender.email_address)
        print(Fore.YELLOW + item.subject)
        print()
        # print(item.to_recipients)
        # print(item.datetime_received)
        letter = item.body
        print(Fore.WHITE + make_soup(letter))


if __name__ == '__main__':
    read_email()
