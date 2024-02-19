import os
import pickle
import os
import io
# Gmail API utils
import urllib
import requests
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
from bs4 import BeautifulSoup
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfReader, PdfWriter, PdfMerger
from reportlab.lib.pagesizes import letter


# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']
our_email = ''


class Owner:
    def __init__(self, name, email, address, arrival, first, departure, cell):
        self.name = name
        self.email = email
        self.address = address
        self.arrival = arrival
        self.first = first
        self.departure = departure
        self.cell = cell

    @classmethod
    def get_name(cls, text):
        sub1 = " for "
        sub2 = "Personal Information"
        idx1 = text.index(sub1)
        idx2 = text.index(sub2)

        res = ''
        for idx in range(idx1 + len(sub1), idx2):
            res = res + text[idx]
        return res

    @classmethod
    def get_cell(cls, text):
        sub1 = "Cell Phone:"
        sub2 = "Email:"
        idx1 = text.index(sub1)
        idx2 = text.index(sub2)

        res = ''
        for idx in range(idx1 + len(sub1), idx2):
            res = res + text[idx]
        cell = ' '.join(res.split())
        return cell

    @classmethod
    def get_address(cls, text):
        sub1 = "Address:"
        sub2 = "Cell Phone"
        idx1 = text.index(sub1)
        idx2 = text.index(sub2)

        res = ''
        for idx in range(idx1 + len(sub1), idx2):
            res = res + text[idx]
        address = ' '.join(res.split())
        return address

    @classmethod
    def get_email(cls, text):
        sub1 = "Email:"
        sub2 = "Arrival and"
        idx1 = text.index(sub1)
        idx2 = text.index(sub2)

        res = ''
        for idx in range(idx1 + len(sub1), idx2):
            res = res + text[idx]
        return res

    @classmethod
    def get_arrival(cls, text):
        sub1 = "Arrival Date:"
        sub2 = "First Ski"
        idx1 = text.index(sub1)
        idx2 = text.index(sub2)

        res = ''
        for idx in range(idx1 + len(sub1), idx2):
            res = res + text[idx]
        return res

    @classmethod
    def get_first(cls, text):
        sub1 = "First Ski:"
        sub2 = "Last Ski"
        idx1 = text.index(sub1)
        idx2 = text.index(sub2)

        res = ''
        for idx in range(idx1 + len(sub1), idx2):
            res = res + text[idx]
        return res

    @classmethod
    def get_last(cls, text):
        sub1 = "Last Ski Date:"
        sub2 = "Skier Information"
        idx1 = text.index(sub1)
        idx2 = text.index(sub2)

        res = ''
        for idx in range(idx1 + len(sub1), idx2):
            res = res + text[idx]
        return res

    @classmethod
    def create_owner(cls, text):
        new_owner = Owner(Owner.get_name(text), Owner.get_email(text), Owner.get_address(text), Owner.get_arrival(text),
                         Owner.get_first(text), Owner.get_last(text), Owner.get_cell(text))
        return new_owner


class Skier:
    def __init__(self, name, height, weight, ski_type, age, gender, package, ski, boot, insurance):
        self.name = name
        self.height = height
        self.weight = weight
        self.ski_type = ski_type
        self.age = age
        self.gender = gender
        self.package = package
        self.ski = ski
        self.boot = boot
        self.insurance = insurance

    @classmethod
    def get_name(cls, text, count):

        skiers = []
        while count > 0:
            sub1 = 'Skier ' + str(count) + ' ❄❄'
            sub2 = 'Skier ' + str(count) + ' Height:'
            idx1 = text.index(sub1)
            idx2 = text.index(sub2)

            res = ''
            for idx in range(idx1 + len(sub1), idx2):
                res = res + text[idx]
            count = count - 1
            skiers.append(res)
        return skiers

    @classmethod
    def count_skiers(cls, text):
        count = 1
        if text.find('Skier 2') > 0:
            count = 2
        if text.find('Skier 3') > 0:
            count = 3
        if text.find('Skier 4') > 0:
            count = 4
        if text.find('Skier 5') > 0:
            count = 5
        if text.find('Skier 6') > 0:
            count = 6
        return count

    @classmethod
    def get_height(cls, text, count):
        skiers = []
        while count > 0:
            sub1 = 'Skier ' + str(count) + ' Height:'
            sub2 = 'Skier ' + str(count) + ' Weight:'
            idx1 = text.index(sub1)
            idx2 = text.index(sub2)

            res = ''
            for idx in range(idx1 + len(sub1), idx2):
                res = res + text[idx]
            count = count - 1
            skiers.append(res)
        return skiers

    @classmethod
    def get_weight(cls, text, count):
        skiers = []
        while count > 0:
            sub1 = 'Skier ' + str(count) + ' Weight:'
            sub2 = 'Skier ' + str(count) + ' Type:'
            idx1 = text.index(sub1)
            idx2 = text.index(sub2)

            res = ''
            for idx in range(idx1 + len(sub1), idx2):
                res = res + text[idx]
            count = count - 1
            skiers.append(res)
        return skiers

    @classmethod
    def get_type(cls, text, count):
        skiers = []

        while count > 0:
            coordinate = 0
            sub1 = 'Skier ' + str(count) + ' Type:'
            sub2 = 'Skier ' + str(count) + ' Age:'
            idx1 = text.index(sub1)
            idx2 = text.index(sub2)

            res = ''
            for idx in range(idx1 + len(sub1), idx2):
                res = res + text[idx]
            count = count - 1

            if res.strip() == "Beginner":
                coordinate = 147
            if res.strip() == "Intermediate":
                coordinate = 180
            if res.strip() == "Advanced":
                coordinate = 215
            skiers.append(coordinate)
        return skiers

    @classmethod
    def get_age(cls, text, count):
        skiers = []
        while count > 0:
            sub1 = 'Skier ' + str(count) + ' Age:'
            sub2 = 'Skier ' + str(count) + ' Gender:'
            idx1 = text.index(sub1)
            idx2 = text.index(sub2)

            res = ''
            for idx in range(idx1 + len(sub1), idx2):
                res = res + text[idx]
            count = count - 1
            skiers.append(res)
        return skiers

    @classmethod
    def get_gender(cls, text, count):
        skiers = []
        while count > 0:
            coordinate = 0
            sub1 = 'Skier ' + str(count) + ' Gender:'
            sub2 = 'Skier ' + str(count) + ' Package'
            idx1 = text.index(sub1)
            idx2 = text.index(sub2)

            res = ''
            for idx in range(idx1 + len(sub1), idx2):
                res = res + text[idx]
            count = count - 1
            if res.strip() == "Male":
                coordinate = 330
            if res.strip() == "Female":
                coordinate = 360

            skiers.append(coordinate)
        return skiers

    @classmethod
    def get_package(cls, text, count):
        skiers = []
        while count > 0:
            sub1 = 'Skier ' + str(count) + ' Package Preference:'
            sub2 = 'Skier ' + str(count) + ' Ski Preference:'
            idx1 = text.index(sub1)
            idx2 = text.index(sub2)

            res = ''
            for idx in range(idx1 + len(sub1), idx2):
                res = res + text[idx]
            count = count - 1
            skiers.append(res)
        return skiers

    @classmethod
    def get_ski(cls, text, count):
        skiers = []
        while count > 0:
            sub1 = 'Skier ' + str(count) + ' Ski Preference:'
            sub2 = 'Skier ' + str(count) + ' Approximate'
            idx1 = text.index(sub1)
            idx2 = text.index(sub2)

            res = ''
            for idx in range(idx1 + len(sub1), idx2):
                res = res + text[idx]
            count = count - 1
            skiers.append(res)
        return skiers

    @classmethod
    def get_length(cls, text, count):
        skiers = []
        while count > 0:
            sub1 = 'Skier ' + str(count) + ' Approximate Ski/ Board Length:'
            sub2 = 'Skier ' + str(count) + ' Boot Size'
            idx1 = text.index(sub1)
            idx2 = text.index(sub2)

            res = ''
            for idx in range(idx1 + len(sub1), idx2):
                res = res + text[idx]
            count = count - 1
            skiers.append(res)
        return skiers

    @classmethod
    def get_boot(cls, text, count):
        skiers = []
        while count > 0:
            sub1 = 'Skier ' + str(count) + ' Boot Size:'
            sub2 = '--Damage'
            idx1 = text.index(sub1)
            idx2 = text.index(sub2)

            res = ''
            for idx in range(idx1 + len(sub1), idx2):
                res = res + text[idx]
            count = count - 1
            if res:
                skiers.append(res)
            else:
                skiers.append("BOOT")
        return skiers

    @classmethod
    def get_insurance(cls, text, count):
        skiers = []
        while count > 0:
            sub1 = 'Damage Waiver: '
            sub2 = '-'
            idx1 = text.index(sub1)
            idx2 = text.index(sub2)

            res = ''
            for idx in range(idx1 + len(sub1), idx2):
                res = res + text[idx]
            count = count - 1
            skiers.append(res)

        return skiers

    @classmethod
    def get_all(cls, count, name, height, weight, ski_type, age, gender, package, ski, boot, insurance):
        skiers = []

        while count > 0:
            skier = Skier(name[count - 1], height[count - 1], weight[count - 1], ski_type[count - 1], age[count - 1],
                          gender[count - 1], package[count - 1], ski[count - 1], boot[count - 1], insurance[count - 1])
            skiers.append(skier)
            count = count - 1
        return skiers

    @classmethod
    def skier_list(cls, text):
        count = Skier.count_skiers(text)
        skiers = Skier.get_all(count, Skier.get_name(text, count), Skier.get_height(text, count),
                               Skier.get_weight(text, count),
                               Skier.get_type(text, count), Skier.get_age(text, count), Skier.get_gender(text, count),
                               Skier.get_package(text, count), Skier.get_ski(text, count), Skier.get_boot(text, count),
                               Skier.get_insurance(text, count))
        return skiers


def gmail_authenticate():
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)


# get the Gmail API service
service = gmail_authenticate()


# utility functions
def get_size_format(b, factor=1024, suffix="B"):
    """
    Scale bytes to its proper byte format
    e.g:
        1253656 => '1.20MB'
        1253656678 => '1.17GB'
    """
    for unit in ["", "K", "M", "G", "T", "P", "E", "Z"]:
        if b < factor:
            return f"{b:.2f}{unit}{suffix}"
        b /= factor
    return f"{b:.2f}Y{suffix}"


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)


def parse_parts(service, parts, folder_name, message):
    """
    Utility function that parses the content of an email partition
    """
    if parts:
        for part in parts:
            filename = part.get("filename")
            mimeType = part.get("mimeType")
            body = part.get("body")
            data = body.get("data")
            file_size = body.get("size")
            part_headers = part.get("headers")
            if part.get("parts"):
                # recursively call this function when we see that a part
                # has parts inside
                parse_parts(service, part.get("parts"), folder_name, message)
            if mimeType == "text/plain":
                # if the email part is text plain
                if data:
                    text = urlsafe_b64decode(data).decode()
                    # print(text)

            elif mimeType == "text/html":
                soup = ""
                # if the email part is an HTML content
                # save the HTML file and optionally open it in the browser
                if not filename:
                    filename = "index.html"
                filepath = os.path.join(folder_name, filename)
                print("Saving HTML to", filepath)
                with open(filepath, "wb") as f:
                    f.write(urlsafe_b64decode(data))

                file = open(filepath, "r", encoding="utf8")

                soup = BeautifulSoup(file, "lxml")
                text = soup.get_text()
                current_owner = Owner.create_owner(text)
                complete_skiers = Skier.skier_list(text)

                # Everything below needs to be in its own method
                # TODO Rewrite into separate method
                # begin pdf creation
                #
                #
                packet = io.BytesIO()
                my_canvas = canvas.Canvas(packet, pagesize=letter)
                my_canvas.drawString(70, 700, current_owner.name)
                my_canvas.drawString(75, 630, current_owner.email.strip())
                my_canvas.drawString(78, 650, current_owner.cell)
                my_canvas.drawString(70, 681, current_owner.address)
                my_canvas.drawString(252, 740, current_owner.arrival.strip())
                my_canvas.drawString(375, 740, current_owner.first.strip())
                my_canvas.drawString(479, 738, current_owner.departure.strip())

                # begin skiers
                total = len(complete_skiers)
                y1 = 600
                y2 = 575
                count = 0
                while total > 0:
                    if count >= 3:
                        my_canvas.showPage()
                        y1 = 600
                        y2 = 575
                        count = 0
                    if count >= 6:
                        my_canvas.showPage()
                        y1 = 600
                        y2 = 575
                        count = 0

                    my_canvas.drawString(100, y1, complete_skiers[total - 1].name)
                    my_canvas.drawString(308, y1, complete_skiers[total - 1].height)
                    my_canvas.drawString(445, y1, complete_skiers[total - 1].weight)
                    my_canvas.drawString(575, y1, complete_skiers[total - 1].boot)
                    my_canvas.rect(complete_skiers[total - 1].ski_type - count, y2 - 4, 20, 17)
                    my_canvas.drawString(298, y2, complete_skiers[total - 1].age)
                    my_canvas.rect(complete_skiers[total - 1].gender, y2 - 4, 30, 17)
                    my_canvas.drawString(400, y2, complete_skiers[total - 1].insurance)
                    total = total - 1
                    count = count + 1
                    if count == 1:
                        y1 = y1 - 155
                        y2 = y2 - 155
                    if count == 2:
                        y1 = y1 - 153
                        y2 = y2 - 153
                    if count == 3:
                        y1 = y1 - 200
                        y2 = y2 - 200

                my_canvas.save()
                # move to the beginning of the StringIO buffer
                packet.seek(0)
                # create a new PDF with Reportlab
                new_pdf = PdfReader(packet)
                # read your existing PDF
                # existing_pdf = PdfReader(open("RentalCard.pdf", "rb"))
                output = PdfWriter()
                # add the "watermark" (which is the new pdf) on the existing page

                if len(complete_skiers) > 3:
                    existing_pdf = PdfReader(open("RentalCard2.pdf", "rb"))
                elif len(complete_skiers) > 6:
                    existing_pdf = PdfReader(open("RentalCard3.pdf", "rb"))
                else:
                    existing_pdf = PdfReader(open("RentalCard.pdf", "rb"))

                for i in range(len(existing_pdf.pages)):
                    page = existing_pdf.pages[i]
                    page.merge_page(new_pdf.pages[i])
                    output.add_page(page)

                # finally, write "output" to a real file
                savepath = os.path.join(folder_name, current_owner.name + ".pdf")
                output_stream = open(savepath, "wb")
                output.write(output_stream)

                output_stream.close()
                print(text)


def read_message(service, message):
    """
    This function takes Gmail API `service` and the given `message_id` and does the following:
        - Downloads the content of the email
        - Prints email basic information (To, From, Subject & Date) and plain/text parts
        - Creates a folder for each email based on the subject
        - Downloads text/html content (if available) and saves it under the folder created as index.html
        - Downloads any file that is attached to the email and saves it in the folder created
    """
    msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
    # parts can be the message body, or attachments
    payload = msg['payload']
    headers = payload.get("headers")
    parts = payload.get("parts")
    folder_name = "Reservations"
    has_subject = False

    parse_parts(service, parts, folder_name, message)
    print("=" * 50)


def search_messages(service, query):
    result = service.users().messages().list(userId='me', q=query + ' is:unread').execute()
    messages = []
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId='me', q=query, pageToken=page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
    return messages


# get emails that match the query you specify
results = search_messages(service, "Reservation")
# for each email matched, read it (output plain/text to console & save HTML and attachments)
for message in results:
    read_message(service, message)

# Skier 1 Approximate Ski/ Board Length: Skier 1 Boot Size: 9.5 mens--Damage Waiver: Yes (recommended)--❄❄
# Skier 2 Approximate Ski/ Board Length: Skier 2 Boot Size: 9 women's wide--Damage Waiver: Yes (recommended)--❄❄
# Skier 1 Approximate Ski/ Board Length: Skier 1 Boot Size: --Damage Waiver: Yes (recommended)--❄❄
# Skier 1 Approximate Ski/ Board Length: Skier 1 Boot Size: 8--Damage Waiver: Yes (recommended)--❄❄
# Skier 2 Approximate Ski/ Board Length: Skier 2 Boot Size: 11--Damage Waiver: Yes (recommended)--❄❄