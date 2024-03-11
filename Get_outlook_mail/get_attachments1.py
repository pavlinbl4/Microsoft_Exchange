import os.path
from exchangelib import FileAttachment, ItemAttachment, Message
from colorama import Fore
from read_outlook import authorization, make_soup
from pathlib import Path


def messages_history():
    with open('received_messages.txt', 'r') as messages_history_file:
        received_messages_id = messages_history_file.readlines()
    return [x.strip() for x in received_messages_id]


def download_attachments():
    account = authorization()

    received_messages_id = messages_history()

    for item in account.inbox.all().order_by('-datetime_received')[:5]:

        if item.message_id not in received_messages_id:
            print("*" * 100)
            print(Fore.RED + item.sender.name)
            print(Fore.GREEN + item.sender.email_address)
            print(Fore.YELLOW + item.subject)
            print()
            letter = item.body
            print(Fore.WHITE + make_soup(letter))
            for attachment in item.attachments:
                if isinstance(attachment, FileAttachment):
                    local_path = os.path.join(f"{Path.home()}/Downloads/Ъ downloads", attachment.name)
                    with open(local_path, "wb") as f:
                        f.write(attachment.content)
                    print("Saved attachment to", local_path)
                elif isinstance(attachment, ItemAttachment):
                    if isinstance(attachment.item, Message):
                        print(attachment.item.subject, attachment.item.body)
            with open('received_messages.txt', 'a') as messages_history_file:
                messages_history_file.write(f'{item.message_id}\n')


if __name__ == '__main__':
    download_attachments()
