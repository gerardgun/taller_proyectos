import psycopg2
from psycopg2 import Error


def test_connection():
    try:
        # Connect to an existing database
        connection = psycopg2.connect(user="modulo4",
                                      password="modulo4",
                                      host="128.199.1.222",
                                      port="5432",
                                      database="delati")

        cursor = connection.cursor()
        print("PostgreSQL server information")
        print(connection.get_dsn_parameters(), "\n")
        cursor.execute("SELECT version();")
        record = cursor.fetchone()
        print("You are connected to - ", record, "\n")

    except (Exception, Error) as error:
        print("Error while connecting to PostgreSQL", error)
    finally:
        if (connection):
            cursor.close()
            connection.close()
            print("PostgreSQL connection is closed")