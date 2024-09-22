import win32com.client
import pandas as pd

# Подключение к Active Directory
ad = win32com.client.Dispatch("ADODB.Connection")
ad.Open("Provider=ADsDSOObject;")

# Чтение данных из Excel-файла
df = pd.read_excel("users.xlsx")

# Итерация по строкам в DataFrame
for index, row in df.iterrows():
    try:
        # Получение данных из строки
        username = row["username"]
        password = row["password"]
        firstname = row["firstname"]
        lastname = row["lastname"]
        email = row["email"]

        # Создание нового объекта пользователя
        user = ad.CreateObject("ADSystemInfo").GetObject("LDAP://CN=" + username + ",OU=Users,DC=domain,DC=com")

        # Задание атрибутов пользователя
        user.SetInfo("sAMAccountName", username)
        user.SetInfo("userPrincipalName", username + "@domain.com")
        user.SetInfo("displayName", firstname + " " + lastname)
        user.SetInfo("mail", email)
        user.SetInfo("userPassword", password)

        # Создание пользователя в Active Directory
        user.SetInfo("objectClass", "user")
        user.SetInfo("objectCategory", "person")
        user.SetInfo("userAccountControl", 512) # Включить учетную запись
        user.SetInfo("givenName", firstname)
        user.SetInfo("sn", lastname)

        # Сохранение изменений
        user.SetInfo()

        print(f"Пользователь {username} успешно создан.")

    except Exception as e:
        print(f"Ошибка при создании пользователя {username}: {e}")

# Закрытие соединения с Active Directory
ad.Close()