import mariadb
import sys
import pandas as pd

try:
    conn = mariadb.connect(
        user="****",
        password="****",
        host="****",
        port=****,
        database="fiction"

    )
except mariadb.Error as e:
    print(f"Error connecting to MariaDB Platform: {e}")
    sys.exit(1)

cur = conn.cursor()


# Insert
def insert():
    # Values to insert
    title = input("Title: ").strip()
    author = input("Author: ").strip()
    fandom = input("Fandom: ").split(", ")  # Separate fandoms by ','
    while True:
        fandom_type = input("Type - F for Fanfic, W for web, Q for quest: ").lower()
        if fandom_type == 'f':
            fandom_type = "Fanfiction"
            break
        elif fandom_type == 'w':
            fandom_type = "Web Fiction"
            break
        elif fandom_type == "q":
            fandom_type = "Quest"
            break
    url = input("URL: ").strip()
    word_count = int(input("Word Count: "))
    chapters = int(input("Chapter Count: "))
    rating = int(input("Rating from 0-10: "))
    series = input("Series: ").strip()
    while True:
        status = input("Status - O for ongoing, C for completed: ").lower()
        if status == 'o':
            status = "Ongoing"
            break
        elif status == 'c':
            status = "Completed"
            break
    published = input("Date Published: ")
    updated = input("Date Updated: ")

    sql = "INSERT INTO works(title, author, type, url, word_count, chapters, rating, series, status, published, " \
          "updated) VALUES( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
    data = (title, author, fandom_type, url, word_count, chapters, rating, series, status, published, updated)

    cur.execute(sql, data)
    conn.commit()

    sql = "SELECT fic_id FROM works WHERE title=? AND author=?"
    data = (title, author)
    cur.execute(sql, data)

    fic_id = cur.fetchone()[0]

    for item in fandom:
        sql = "INSERT INTO fandom(fic_id, fandom_name) VALUES (?, ?)"
        data = (fic_id, item)
        cur.execute(sql, data)

    conn.commit()


# Update
def update():
    sql = "SELECT works.fic_id AS ID, title AS Title, author AS Author, " \
          "GROUP_CONCAT(fandom.fandom_name SEPARATOR ', ') AS Fandom, " \
          "type AS Type, url AS URL, FORMAT(word_count, N'N0') AS 'Word Count', chapters AS Chapters, " \
          "FORMAT(CAST(word_count / chapters AS INT), N'N0') AS 'Avg words'," \
          "rating AS Rating, series as Series, status as Status, published as 'Published Date', " \
          "updated AS 'Updated Date' FROM works " \
          "LEFT JOIN fandom ON (works.fic_id = fandom.fic_id) " \
          "group by fandom.fic_id " \
          "order by works.fic_id "

    cur.execute(sql)

    # Prints result of SQL statement to file
    columns = [desc[0] for desc in cur.description]
    data = cur.fetchall()
    df = pd.DataFrame(list(data), columns=columns)

    writer = pd.ExcelWriter('Fictions.xlsx')  # Name of file
    df.to_excel(writer, sheet_name='Fictions', index=False)
    writer.save()
    print("File saved.")


while True:
    option = input("1 to add a new option. 2 to update excel file. "
                   "> ")
    if option == '1':
        insert()
    elif option == '2':
        update()
        break

# Close connection
cur.close()
conn.close()
