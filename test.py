
import smtplib
server = smtplib.SMTP('smtp.office365.com', 587, timeout=10)
server.starttls()
try:
    server.login('xuecz1@lenovo.com', 'Zhang2000')
    print('认证成功！')
except smtplib.SMTPAuthenticationError as e:
    print(f'认证失败: {e}')
finally:
    server.quit()
