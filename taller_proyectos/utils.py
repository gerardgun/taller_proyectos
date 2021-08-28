import psycopg2
import re
import unidecode
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


def get_data():
    connection = psycopg2.connect(user="modulo4",
                                  password="modulo4",
                                  host="128.199.1.222",
                                  port="5432",
                                  database="delati")

    cursor = connection.cursor()
    cursor.execute("""select od.id_ofertadetalle,
            od.descripcion_normalizada
            from oferta o
            inner join oferta_detalle od
            on (o.id_oferta=od.id_oferta)
            where  o.id_estado is null and od.ofertaperfil_id=6
            order by id_ofertadetalle;"""
           )
    records = cursor.fetchall()
    data = []
    for record in records:
        # delete tilde
        text = unidecode.unidecode(record[1].strip())
        data.append(text.strip())

    for dato in data:
        print(dato)
    #
    #     text = re.sub("[^a-zA-Z0-9]+", "", record[1])
    # print("You are connected to - ", record, "\n")
