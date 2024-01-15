import os.path
from exchangelib import FileAttachment, ItemAttachment, Message
from colorama import Fore

from read_outlook import authorization, make_soup


def download_attachments():
    account = authorization()

    for item in account.inbox.all().order_by('-datetime_received')[:5]:
        print("*" * 100)
        print()
        print(Fore.RED + item.sender.name)
        print(Fore.GREEN + item.sender.email_address)
        print(Fore.YELLOW + item.subject)
        print()
        letter = item.body
        print(Fore.WHITE + make_soup(letter))
        for attachment in item.attachments:
            if isinstance(attachment, FileAttachment):
                local_path = os.path.join("/Volumes/big4photo/Downloads/Ъ downloads", attachment.name)
                with open(local_path, "wb") as f:
                    f.write(attachment.content)
                print("Saved attachment to", local_path)
            elif isinstance(attachment, ItemAttachment):
                if isinstance(attachment.item, Message):
                    print(attachment.item.subject, attachment.item.body)


if __name__ == '__main__':
    download_attachments()
