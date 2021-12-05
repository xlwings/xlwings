import xlwings as xw


def main2():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello2 xlwings!":
        sheet["A1"].value = "Bye2 xlwings!"
    else:
        sheet["A1"].value = "Hello2 xlwings!"


@xw.func
def hello2(name):
    return f"Hello2 {name}!"


if __name__ == "__main__":
    xw.Book("permission.xlsm").set_mock_caller()
    main2()
