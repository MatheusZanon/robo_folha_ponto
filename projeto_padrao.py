from components.importacao_automacao_web import auto_web
from components.importacao_automacao_excel import auto_excel

driver, sleep, By = auto_web()

driver.get("https://www.selenium.dev/selenium/web/web-form.html")

sleep(2)

text_box = driver.find_element(By.NAME, value="my-text")
text_box.send_keys('teste do selenium asimov acadimi')
submit_button = driver.find_element(By.CSS_SELECTOR, value="button")
submit_button.click()

message = driver.find_element(By.ID, value="message")
print('Mensagem final', message.text)

sleep(2)