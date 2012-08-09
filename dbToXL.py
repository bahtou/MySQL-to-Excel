import MySQLdb
import logging
import sys


# initialize logging
log = logging.getLogger("dataNitro")
log.setLevel(logging.INFO)
fh = logging.FileHandler("dataNitro.log")
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
log.addHandler(fh)


# MySQL connection
def MySQLConn(host="localhost", user='', pwd='', db=''):
    try:
        conn = MySQLdb.connect(
                    host=host,
                    user=user,
                    passwd=pwd,
                    db=db)
        log.info('Connected to MySQL')
        cursor = conn.cursor(MySQLdb.cursors.DictCursor)
    except MySQLdb.Error, e:
        print "Error %d: %s" % (e.args[0], e.args[1])
        sys.exit(1)
    return conn, cursor


# label header in Excel
def header(head):
    CellRange("A1:D1").value = head
    log.info('header placed')


# insert data into cells
def cellInsert(cursor, sql):
    cursor.execute(sql)
    results = cursor.fetchall()
    for i, row in enumerate(results):
        CellRange("A%s:D%s" % (i + 2, i + 2)).value = row['postsBy'], row['headLine'], row['headLineLink'], row['postTime']
    autofit()
    log.info('Done inserting into MySQL')


# close db connection
def hangUp(cursor, conn):
    cursor.close()
    conn.close()
    log.info('Disconnceted from MySQL')


# start file execution
def main():
    clear_sheet()
    conn, cursor = MySQLConn(user="user", pwd="password", db="techcrunch")

    head = ['postsBy', 'headLine', 'headLineLink', 'postTime']
    header(head)

    sql = "SELECT postsBy, headLine, headLineLink, postTime FROM homepage"
    cellInsert(cursor, sql)

    hangUp(cursor, conn)

# EXECUTE
if __name__ == "__main__":
    main()
