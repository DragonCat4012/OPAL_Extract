END = "\033[0m"
FAIL = "\033[91m"


class Logger:
    @staticmethod
    def info(value: str, tabs: int = 0):
        pre = "".join(["\t" for _ in range(tabs)])
        print(pre + value)

    @staticmethod
    def error(value: str, tabs: int = 0):
        pre = "".join(["\t" for _ in range(tabs)])
        print(FAIL + pre + "⚠️ " + value + END)
