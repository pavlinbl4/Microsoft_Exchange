from exchangelib import Message, Mailbox

from get_mail.read_outlook import authorization

m = Message(
    account=authorization(),
    subject="Python forever 6",
    body="Text in the body",
    to_recipients=[
        Mailbox(email_address="pavlenko.evgeniy@gmail.com"),
        # Mailbox(email_address="pavlinbl4@gmail.com"),
    ],
    #Simple strings work, too additional copy
    # cc_recipients=["calvoegrasso@gmail.com", "x79046309437@gmail.com"],
    # bcc_recipients=[
    #     Mailbox(email_address="tel.89046309437@gmail.com"),
    #     "stockphoto.avtor@gmail.com",
    # ],
    reply_to=['epavlenko@kommersant.ru']




)

m.send()
