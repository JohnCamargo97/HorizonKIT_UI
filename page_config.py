butON = "QPushButton {\n   color: rgb(25, 25, 25);\n   border-top-left-radius: 10px;\n   border-top-right-radius: 10px;\n   background-color: rgb(240, 240, 240);\n   border: 0px solid;\n }\n QPushButton:pressed {\n \n color: rgb(25, 25, 25); \n background-color: rgb(230, 230, 230);\n \n }"
butOFF = "QPushButton {\n   color: rgb(230, 230, 230);\n   border-top-left-radius: 10px;\n   border-top-right-radius: 10px;\n   background-color: rgb(115, 115, 115);\n   border: 0px solid;\n }\n QPushButton:pressed {\n \n color: rgb(25, 25, 25); \n background-color: rgb(230, 230, 230);\n \n }"

class FUNCTIONS():
    def toggle_page(state):
        if state == True:
            select = butON
            return select
        else:
            select = butOFF
            return select 