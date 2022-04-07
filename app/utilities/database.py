# Created by Victor Stanescu at 10/14/2019

# Enter feature description here

# Enter steps here

from psycopg2 import pool



class Database:
    __connection_pool = None

    @classmethod
    def initialise(cls, host="dbrobotscluster.cluster-cmtwo2g5wgkl.eu-west-1.rds.amazonaws.com",
                 database="uipathdb",
                 user="dev_stanv",
                 password="lpRPA1",
                 minconn=2,
                 maxconn=10):

        cls.__connection_pool = pool.SimpleConnectionPool(minconn=minconn,
                                                          maxconn=maxconn,
                                                          database=database,
                                                          user=user,
                                                          password=password,
                                                          host=host)
    @classmethod
    def get_connection(cls):
        return cls.__connection_pool.getconn()

    @classmethod
    def return_connection(cls, connection):
        cls.__connection_pool.putconn(connection)

    @classmethod
    def close_all_connections(cls):
        cls.__connection_pool.closeall()



class CursorFromConnectionFromPool:
    def __init__(self):
        self.connection = None
        self.cursor = None

    def __enter__(self):
        self.connection = Database.get_connection()
        self.cursor = self.connection.cursor()
        return self.cursor

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_val is not None:
            self.connection.rollback()
        else:
            self.cursor.close()
            self.connection.commit()
        Database.return_connection(self.connection)
