from flask import Flask, render_template, request
from openpyxl import load_workbook
from sqlalchemy.orm.session import sessionmaker
from database import engine, Book
from sqlalchemy.orm import scoped_session

app = Flask(__name__)

db = scoped_session(sessionmaker(bind = engine)) #лучше использовать внутри функции

@app.route("/")
def homepage():
    return render_template("index.html")


@app.route("/us/")
def info():
    return render_template("aboutus.html")

@app.route("/books/")
def books():
    # v_2
    # excel = load_workbook("tales.xlsx")
    # page = excel["Sheet"]

    # object_list = [[tale.value, tale.offset(column = 1).value] for tale in page["A"][1:]]
    # return render_template("books.html", object_list = object_list)

    # v_3
    # session = sessionmaker(engine)()
    # books = session.query(Book)
    # session.commit()

    with engine.connect() as con:
        books = con.execute("""SELECT * FROM "Book";""")
        
    return render_template("books_table.html", object_list = books)

@app.route("/authors/")
def authors():
    # v_1
    # excel = load_workbook("tales.xlsx")
    # page = excel["Sheet"]
    # authors = {author.value for author in page["B"][1:]}

    with engine.connect() as con:
        authors = con.execute("""SELECT DISTINCT author FROM "Book";""")

    return render_template(
        "authors.html", authors = list(authors)
    )

@app.route("/add/", methods = ["POST"])
def add():
    f = request.form
    # v_1
    # excel = load_workbook("tales.xlsx")
    # page = excel["Sheet"]
    # last = len(page["A"]) + 1
    # page[f"A{last}"] = f["book"]
    # page[f"B{last}"] = f["author"]
    # excel.save("tales.xlsx")
    book = f["book"]
    author = f["author"]
    # url = f["url"]
    ids = db.execute('SELECT id FROM "Book" ORDER BY id DESC;')
    max_id = ids.first().id
    c_id = max_id + 1

    with engine.connect() as con:
        add = con.execute(f"""
        INSERT INTO "Book" (id, name, author) 
        VALUES ({c_id}, '{book}', '{author}');""")
    return "Form Accepted!"

@app.route("/book/<num>/")
def book(num):
    excel = load_workbook("tales.xlsx")
    page = excel["Sheet"]
    object_list = [[tale.value, tale.offset(column = 1).value, tale.offset(column = 2).value] for tale in page["A"][1:]]
    obj = object_list[int(num)]
    obj.append(num)
    return render_template("book.html", obj = obj)

@app.route("/book/<num>/edit/")
def book_edit(num):
    num = int(num) + 2
    excel_file = load_workbook("tales.xlsx")
    page = excel_file["Sheet"]
    tale = page[f"A{num}"]
    author = page[f"B{num}"]
    image = page[f"C{num}"]
    obj = [tale.value, author.value, image.value, num]
    return render_template("book_edit.html", obj = obj)

@app.route("/book/<num>/save/", methods = ["POST"])
def book_save(num):
    num = int(num)
    excel_file = load_workbook("tales.xlsx")
    page = excel_file["Sheet"]
    form = request.form
    page[f"A{num}"] = form["tale"]
    page[f"B{num}"] = form["author"]
    page[f"C{num}"] = form["image"]
    excel_file.save("tales.xlsx")
    return "Saved!"