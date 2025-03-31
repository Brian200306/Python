import mysql.connector

def get_connection():
    try:
        connection = mysql.connector.connect(
            host='192.168.0.109',
            user='root',
            password='Emp3698521',
            database='empral'
        )
        return connection
    except mysql.connector.Error as err:
        print("Erro no banco de dados:", err)
        return None