class Team:
    member: list[str] = []
    color: str = ""

    def __init__(self, members):
        self.member = [member for member in members if str(member).isdigit()]

    def is_member(self, person):
        """Check if person is in group"""
        return person in self.member

    def __str__(self):
        return " ".join(map(str, self.member))


class SingularPerson:
    name: str
    surname: str
    color = "000000"

    def __init__(self, name, surname):
        print(name, surname)
        self.name = name
        self.surname = surname

    def get_key(self):
        return f"{self.surname.lower()} {self.name.lower()}"
