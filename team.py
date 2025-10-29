class Team:
    member: list[str] = []
    color: str = ""

    def __init__(self, members):
        self.member = members

    def is_member(self, person):
        return person in self.member