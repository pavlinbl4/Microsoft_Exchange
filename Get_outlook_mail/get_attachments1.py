import os.path
from exchangelib import FileAttachment
from colorama import Fore
from read_outlook import authorization, make_soup
from pathlib import Path
from loguru import logger

script_dir = os.path.dirname(os.path.abspath(__file__))
logger.add(f'{script_dir}/log.log', rotation='10:00')


def messages_history() -> list:
    with open(f'{script_dir}/received_messages.txt', 'r') as messages_history_file:
        received_messages_id = messages_history_file.readlines()
    return [x.strip() for x in received_messages_id]


def download_attachments():
    account = authorization()

    received_messages_id = messages_history()
    # received_messages_id = []

    for item in account.inbox.all().order_by('-datetime_received')[:5]:
        logger.info(f'{item.sender.email_address = }')

        if item.message_id not in received_messages_id:

            if len(item.attachments) > 0:
                print("*" * 100)
                print(Fore.RED + item.sender.name)
                print(Fore.GREEN + item.sender.email_address)
                print(Fore.YELLOW + item.subject)
                print()
                letter = item.body
                print(Fore.WHITE + make_soup(letter))
                for attachment in item.attachments:
                    if isinstance(attachment, FileAttachment) and attachment.size > 9000:
                        local_path = os.path.join(f"{Path.home()}/Downloads/Ъ downloads", attachment.name)
                        with open(local_path, "wb") as f:
                            f.write(attachment.content)
                        logger.info(f"Saved attachment to {local_path}")
                with open(f'{script_dir}/received_messages.txt', 'a') as messages_history_file:
                    messages_history_file.write(f'{item.message_id}\n')
        else:
            logger.info('No new messages')


if __name__ == '__main__':
    download_attachments()
