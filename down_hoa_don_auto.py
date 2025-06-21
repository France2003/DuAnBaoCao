# import thư viện
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
# Gửi email
import smtplib
import shutil
#"Mã hóa kết nối với gmail"
import ssl
from email.message import EmailMessage
ma_tra_cuu_list =[
    "PZH_FWQ4BN3",
    "B1HEIRR8N0WP",
    "VBHKSL682918"
]
email_nhan = "phapcv2003@gmail.com"
email_gui = "phapcv2003@gmail.com"
password_connect_email = "xvcj lrok jtda kpmt"
name_folder_download = "HoaDonDienTu"
# Hàm tự động mở web v tra cứu hóa đơn
def input_hoa_don_auto(ma_tra_cuu):
    os.makedirs(name_folder_download, exist_ok=True)
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    prefs = {
        "download.default_directory": os.path.abspath(name_folder_download),
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    }
    options.add_experimental_option("prefs", prefs)
    # Mở trang web
    driver = webdriver.Chrome(options=options)
    driver.get("https://www.meinvoice.vn/tra-cuu")
    try:
        # Input nhập mã
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.NAME, "txtCode"))
        )
        input_code_value = "txtCode"
        input_code_element = driver.find_element(By.NAME, input_code_value)
        input_code_element.send_keys(ma_tra_cuu)
        print("Đang tra cứu mã: ", ma_tra_cuu)
        # Button tra cứu mã
        button_search_value = "btnSearchInvoice"
        btn_search_element= WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, button_search_value))
        )
        btn_search_element.click()
        print("Đã nhấn tra cứu")
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Tải hóa đơn')]"))
        )
        print("Hóa đơn đã sẵn sàng để tải")
        return  driver
    except Exception as error:
        print("Mã tra cứu sai", error)
        driver.quit()
        return None
def tai_hoa_don(driver):
    try:
        download_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "download"))
        )
        # Scroll đến đúng vị trí, tránh bị che
        driver.execute_script("arguments[0].scrollIntoView(true);", download_btn)
        time.sleep(1)  # Đợi UI ổn định
        driver.execute_script("arguments[0].click();", download_btn)
        print("Đã nhấn nút chính 'Tải hóa đơn'")

        pdf_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "txt-download-pdf"))
        )
        driver.execute_script("arguments[0].click();", pdf_option)
        print("Đã chọn tải hóa đơn dạng PDF")
        time.sleep(10)  # Chờ file tải
        return True
    except Exception as e:
        print("Lỗi khi tải hóa đơn PDF:", e)
        return False
def search_file_new(folder):
    print("Đang chờ file tải xong...")
    timeout = 20
    start_time = time.time()
    while time.time() - start_time < timeout:
        files = os.listdir(folder)
        full_paths = [os.path.join(folder, f) for f in files if f.endswith(".pdf") and not f.endswith(".crdownload")]
        if full_paths:
            file_path = max(full_paths, key=os.path.getctime)
            print("Đã tìm thấy file:", file_path)
            return file_path
        time.sleep(1)
    print("Không tìm thấy file nào.")
    return None
def gui_file_ve_gmail(file_path, gmail_gui, mat_khau_app, gmail_nhan):
    subject = "Hóa đơn điện tử từ meInvoice"
    body = "Đây là file hóa đơn bạn vừa tra cứu tự động."
    # Tạo nội dung email
    msg = EmailMessage()
    msg["From"] = gmail_gui
    msg["To"] = gmail_nhan
    msg["Subject"] = subject
    msg.set_content(body)
    # Đọc nội dung file đính kèm
    with open(file_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(file_path)
    # Gắn file vào email
    msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
    # Gửi email qua Gmail bằng kết nối SSL
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(gmail_gui, mat_khau_app)
        server.send_message(msg)
        print(f"Đã gửi file '{file_name}' về Gmail:", gmail_nhan)
if __name__ == "__main__":
    print("Bắt đầu tra cứu và gửi hóa đơn...\n")
    for ma_tra_cuu in ma_tra_cuu_list:
        print(f"Đang xử lý mã: {ma_tra_cuu}")
        driver = input_hoa_don_auto(ma_tra_cuu)
        if driver:
            if tai_hoa_don(driver):
                driver.quit()  # Đóng trình duyệt sau khi tải
                file_moi = search_file_new(name_folder_download)
                if file_moi:
                    # Đổi tên file thành đúng mã
                    new_file_path = os.path.join(name_folder_download, f"{ma_tra_cuu}.pdf")
                    shutil.move(file_moi, new_file_path)
                    print("Đã đổi tên file thành:", new_file_path)
                    gui_file_ve_gmail(new_file_path, email_gui, password_connect_email, email_nhan)
                else:
                    print("Không tìm thấy file nào để gửi.")
            else:
                print("Tải hóa đơn thất bại.")
        else:
            print("Không thể mở trình duyệt hoặc nhập mã.")
        print("--------------------------------------------------\n")






