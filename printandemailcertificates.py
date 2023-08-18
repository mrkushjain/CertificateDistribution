import csv
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def send_email(sender_email, sender_password, recipient_email, recipient_name, subject, message, attachment_path):
    # Set up the MIME objects
    msg = MIMEMultipart()
    msg['From'] = f"Kush Jain <{sender_email}>"
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Create the HTML content for the email
    html_content = f"""
    <html>
    <body>
        <p>Hare Krishna {recipient_name},</p>
        <p>{message}</p>
        <p>Your PDF certificate from ISKCON Agra is attached with this email.</p>
    </body>
    </html>
    """

    msg.attach(MIMEText(html_content, 'html'))

    with open(attachment_path, "rb") as attachment:
        part = MIMEApplication(attachment.read(), Name="attachment.pdf")
        part['Content-Disposition'] = f'attachment; filename="{attachment_path}"'
        msg.attach(part)

    # Connect to the Gmail server
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()

    # Login to the Gmail account
    server.login(sender_email, sender_password)

    # Send the email
    server.sendmail(sender_email, recipient_email, msg.as_string())

    # Quit the server
    server.quit()

def convert_image_to_pdf(image_path, output_pdf_path):
    img = Image.open(image_path)
    
    # Calculate the aspect ratio to maintain the image's proportions
    aspect_ratio = img.width / img.height
    
    # Create a new PDF with the calculated size
    c = canvas.Canvas(output_pdf_path, pagesize=(img.width, img.height))
    
    # Draw the image onto the PDF
    c.drawImage(image_path, 0, 0, width=img.width, height=img.height)
    
    # Save the PDF
    c.save()

def add_text_to_image(image_path, text, output_path):
    image = Image.open(image_path)
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype("VLADIMIR.TTF", 30)  # Change the font and size as needed
    # font = ImageFont.truetype("ariel", 24)
    
    x0 = image.width  / 2 + 15
    y0 = image.height / 2 + 55
    bb_l, bb_t, bb_r, bb_b = draw.textbbox((int(x0), int(y0)), text)
    x = x0 + (bb_r - bb_l) / 2
    y = y0 + (bb_b - bb_t) / 2
    draw.text((int(x), int(y)),text, fill="black", font=font)
    
    image.save(output_path)

def process_csv(csv_file):
    image_file = "blank.jpg" 
    with open(csv_file, 'r+', encoding="utf8") as file:
        reader = csv.DictReader(file)
        fieldnames = reader.fieldnames

        if 'status' not in fieldnames:
            fieldnames.append('status')

        rows = list(reader)

        file.seek(0)
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()

        for row in rows:
            status = row['status']
            if status == "Sent":
                continue
            try:

                email = row['email']
                name =row[fieldnames[0]]
                number = row['number']
                output_image_path = f"{number}.png"
                output_pdf_path = f"{number}.pdf"

                add_text_to_image(image_file, name, output_image_path)
                print(f"Image with name '{name}' created: {output_image_path}")
                
                convert_image_to_pdf(output_image_path, output_pdf_path)
                print(f"Image '{output_image_path}' converted to PDF '{output_pdf_path}'")

                sender_email = "YOUR NAME<GMAIL EMAIL>"
                sender_password = "APP PASSWORD"
                recipient_email = email
                recipient_name = name
                subject = f"Hare Krishna {name}. Congratulations"
                message = row['message']
                attachment_path = output_pdf_path
                send_email(sender_email, sender_password, recipient_email, recipient_name, subject, message, attachment_path)
                print(f"Email for {name} sent to email {email}")
                status = "Sent"
            except:
                status = "Error"

            print(f"Processing for {name} and email {email} . Status {status}")
            row['status'] = status
            writer.writerow(row)



if __name__ == "__main__":
    csv_file = "names.csv"  # Replace with your CSV file path

    process_csv(csv_file)
    print("CSV processing complete.")
