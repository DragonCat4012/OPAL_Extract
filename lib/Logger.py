END = "\033[0m"
FAIL = "\033[91m"
CYAN = "\033[96m"


class Logger:
    @staticmethod
    def info(value: str, tabs: int = 0):
        pre = "".join(["\t" for _ in range(tabs)])
        print(pre + value)

    @staticmethod
    def info_colored(value: str, tabs: int = 0):
        pre = "".join(["\t" for _ in range(tabs)])
        print(CYAN + pre + value + END)

    @staticmethod
    def error(value: str, tabs: int = 0):
        pre = "".join(["\t" for _ in range(tabs)])
        print(FAIL + pre + "⚠️ " + value + END)
