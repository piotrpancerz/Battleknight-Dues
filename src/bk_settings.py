server = "28"
excel_old_path = "../dues.xlsx"
update = False # true if update, false if normal loan check
if update :
    txt_dir_path = "../tmpText/"
    excel_new_path = "../tmpText/dues.xlsx"

else :
    txt_dir_path = "../prodText/"
    excel_new_path = "../dues.xlsx"

title_tuple = ('Mieszczanin','Mieszczanka','Pan','Dama','Panna','Dziedzic','Szlachcic','Szlachcianka','Rycerz','Kniaź','Baronet','Baroneta','Baron','Wicehrabia','Margrabia','Margrabina','Książę')
excel_columns = ({"iteration": "A","name": "ID", "width": 15},{"iteration": "B","name": "Nick", "width": 25},{"iteration": "C","name": "Poziom", "width": 10},{"iteration": "D","name": "Poprzedni stan", "width": 15},{"iteration": "E","name": "Posiadany dług", "width": 15},{"iteration": "F","name": "Kwota do zapłaty", "width": 15},{"iteration": "G","name": "Aktualny stan", "width": 15},{"iteration": "H","name": "Wymagany stan", "width": 15},{"iteration": "I","name": "Nowy dług", "width": 15},{"iteration": "J","name": "Na plus", "width": 15},{"iteration": "K","name": "Seria", "width": 15})