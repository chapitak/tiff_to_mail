import os
import win32com.client as win32

# to_mail_address = '수신자 메일주소를 넣고 주석을 해제해주세요'

# 해당 디렉토리 밑에 tif파일을 담은 폴더를 만들고 폴더명을 만들고 싶은 pdf 파일명으로 적어주시면 됩니다

def make_pdfs():
    pdfs_made_name = []
    for dirname, dirnames, filenames in os.walk('.'):
        if '.git' in dirnames:
        # don't go into any .git directories.
            dirnames.remove('.git')
        # print path to all subdirectories first.
        for subdirname in dirnames:
            subpath = os.path.join(dirname, subdirname)
            tif_file_name = get_tif_file_name(subpath)
            os.system('magick convert '+ tif_file_name + ' ' + '"'+ subpath + '\\' + subdirname+'"' + '.pdf')
            pdfs_made_name.append(subdirname)
    return pdfs_made_name

def get_tif_file_name(path_dir):
    file_list = os.listdir(path_dir)
    file_list.sort()

    tif_file_name = []

    for filename in file_list:
        if filename[-3:] == "tif":
            filename = '"'+ path_dir + '\\' + filename+'"'
            tif_file_name.append(filename)        

    formatted_tif_name = ' '.join(tif_file_name)
    return formatted_tif_name
    
def send_mail(pdfs_made_name):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to_mail_address
    mail.Subject = "[스캔]" + ''.join(pdfs_made_name)
    mail.Body = 'Message body'
    mail.HTMLBody = '**님 안녕하십니까.</br> </br>' + '<br>'.join(pdfs_made_name) + '파일 <br><br> 스캔하고 pdf로 변환해서 보내드립니다 <br><br>'+'감사합니다 <br><br> ※본 메일은 프로그램을 통해 발송된 메일입니다. '
    for pdf_path in pdfs_made_name:        
        attachment  = os.getcwd() + '\\' + pdf_path + '\\' + pdf_path + ".pdf"
        mail.Attachments.Add(attachment)
    mail.Send()

if __name__ == "__main__":
    pdfs_made_name = make_pdfs()
    for pdf_name in pdfs_made_name:
        print(pdf_name)
    print("pdf 파일이 생성되었습니다")
    input("엔터를 누르면 생성된 파일로" + to_mail_address + "님께 메일이 발송됩니다. ")
    send_mail(pdfs_made_name)
    input("발송이 완료되었습니다")
    # pdf_name = input("만들 pdf이름을 입력해주세요> ")


    # tif_file_name = get_tif_file_name()
    # os.system('magick convert '+ tif_file_name + ' ' + '"'+pdf_name+'"' + '.pdf')
    # input("엔터를 누르면 " + to_mail_address + "님께 메일이 발송됩니다. ")
    # # send_mail(pdf_name)
    # input("발송이 완료되었습니다")
