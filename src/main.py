# Custom Function Imports
from src.modules.functions import open_the_intranet_website, log_in, download_excel_file, \
fill_form_with_excel_data, collect_results, export_as_pdf, log_out
# Robocorp Imports
from robocorp.tasks import task

@task
def main():
    open_the_intranet_website()
    log_in()
    download_excel_file()
    fill_form_with_excel_data()
    timestamp = collect_results()
    export_as_pdf(timestamp)
    log_out()


if __name__ == "__main__":
    main()