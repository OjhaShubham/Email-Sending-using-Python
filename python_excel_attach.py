import win32com.client as win32

outlook= win32.Dispatch('Outlook.Application')

message=outlook.CreateItem(0)

message.display()

message.To='shubham-geetashankar.ojha@capgemini.com'

message.CC='shubham-geetashankar.ojha@capgemini.com'

message.Subject='pywin32 Testing'

message.Body='This email is send from python script just ignore it I am learning sending mail from python '

html_body = """
    <div>
        <h1 style="font-family: 'Lucida Handwriting'; font-size: 56; font-weight: bold; color: #9eac9c;">
            Happy Birthday!! 
        </h1>
        <span style="font-family: 'Lucida Sans'; font-size: 28; color: #8d395c;">
            Wishing you all the best on your birthday!!
        </span>
    </div><br>
    <div>
        <img src="https://hips.hearstapps.com/hmg-prod.s3.amazonaws.com/images/cute-birthday-instagram-captions-1584723902.jpg" width=50%>
    </div>
    """

message.HTMLBody = html_body

message.Send()

