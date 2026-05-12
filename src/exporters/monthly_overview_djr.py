import logging


def populate_monthly_overview_djr(
    wb,
    records,
    enable_mediaplanning=True,
    channel_id_to_name_map=None,
):
    """
    Placeholder DJR monthly overview exporter.

    For now, this prevents import errors while we build and test
    FULL CALENDAR YT DJR EMEA first.
    """
    sheet_name = "MONTHLY OVERVIEW YT DJR EMEA"

    if sheet_name not in wb.sheetnames:
        logging.warning(f"Sheet '{sheet_name}' not found. Skipping DJR monthly overview.")
        return

    logging.info("DJR monthly overview exporter not implemented yet. Skipping for now.")
    return