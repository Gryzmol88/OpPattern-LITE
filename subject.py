
class Subject:
    """"Klasa określająca dane zajęcia"""

    def __init__(self, name, classroom, date, start_time, end_time):
        self.name = name
        self.classroom = classroom
        self.date = date
        self.start_time = start_time
        self.end_time = end_time

    def __repr__(self):
        #Stworzyć metode, która wyświetla informację o obiekcie.
        return f'DATA: {self.date.date()}. Sala: {self.classroom}. Przedmiot: {self.name}.' \
               f'Godzina rozpoczęcia: {self.start_time}. Godzina zakończenia: {self.end_time}'
