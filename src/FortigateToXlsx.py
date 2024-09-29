import xlsxwriter


class FortigateToXlsx:
    """Here is created and exported the xlsx file with your configuration"""

    MERGE_FORMAT = {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "black",
        "font_color": "yellow",
    }
    BOLD_ALIGN_FORMAT = {"bold": 1, "border": 1, "align": "center", "valign": "vcenter"}
    GREEN_BOLD_ALIGN_FORMAT = {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "green",
    }
    RED_BOLD_ALIGN_FORMAT = {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "red",
    }
    BLUE_BOLD_ALIGN_FORMAT = {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#00CCFF",
    }
    ALIGN_FORMAT = {
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "font_color": "black",
    }
    SUBTITLE_FORMAT = {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#C0C0C0",
        "font_color": "black",
    }

    def __init__(self) -> None:
        self.workbook = xlsxwriter.Workbook("Report_FORTIGATE.xlsx")
        self.lign = 3
        self.column = 2
        self.vdom_name = ""
        self.worksheet = ""
        self.configuration = ""

    def get_format(self, format: str):
        return self.workbook.add_format(format)

    def get_merge_string(self, first_column, last_column):
        return f"{first_column}{self.lign+1}:{last_column}{self.lign+1}"
