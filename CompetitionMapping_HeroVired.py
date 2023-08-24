import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

url = "https://herovired.com/"

try:
    driver = webdriver.Chrome(executable_path = "C:/Users/Gopal.Jha/OneDrive - Times Group/Desktop/selenium/chromedriver.exe")
    driver.maximize_window()
    driver.get(url)

    elements = driver.find_elements(By.CSS_SELECTOR, ".oxy-tab-content.tabs-contents-5627-tab")
    course_name = []
    program_url = []
    program_fee = []
    class_type = []
    for i in range(len(elements)):
        driver.execute_script("arguments[0].scrollIntoView();", elements[i])
        driver.execute_script("arguments[0].click();", elements[i])
        wait = WebDriverWait(driver, 10)

        url = elements[i].find_elements(By.CSS_SELECTOR, ".ct-link.program-card-collab-container")
        for tags in url:
            program_url.append(tags.get_attribute("href"))

        program = elements[i].find_elements(By.CSS_SELECTOR, "span.ct-span")
        for tags in program:
            course_name.append(tags.get_attribute("innerText"))

    for url in program_url:
        driver.get(url)
        program_fee_exists = driver.find_elements(By.CSS_SELECTOR, "#span-8243-352")
        if program_fee_exists:
            program_fee.append(program_fee_exists[0].get_attribute("innerText"))
        else:
            program_fee.append("")

    list = [course_name[i:i + 8] for i in range(0, len(course_name), 8)]

    # Create a pandas DataFrame from the divided list
    df = pd.DataFrame(list, columns=['Feature', 'University', 'Program', 'Program_Name', 'Type', 'Duration',
                                     'Application_Deadline', 'More Info.'])
    df1 = pd.DataFrame(program_fee, columns=['Program_Fee'])

    merged_data = pd.concat([df, df1], axis=1)
    writer = pd.ExcelWriter('HeroVired.xlsx')
    merged_data.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()

except Exception as e:
    print(e)

finally:
    driver.quit()
