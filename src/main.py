import os
import socket
import smtplib
import yaml

import logging.config
import win32com.client

from croniter import croniter
from datetime import datetime, timedelta
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from time import sleep


APP_NAME: str = ""
APP_VERSION: str = ""

CURRENT_DIR: Path = Path(__file__).parent
MACHINE_NAME: str = ""

LOGGER = logging.getLogger(APP_NAME)

TASK_STATE = {0: 'Unknown',
              1: 'Disabled',
              2: 'Queued',
              3: 'Ready',
              4: 'Running'}


def configure_logger():
    with open(os.path.join(CURRENT_DIR, "logging.yaml"), "r") as config_file:
        config_data = yaml.safe_load(config_file.read())

        logging.config.dictConfig(config_data)
        logging.basicConfig(level=logging.NOTSET)


def read_config():
    global APP_NAME, APP_VERSION

    with open(os.path.join(CURRENT_DIR, "config.yaml"), "r") as config_file:
        config_data = yaml.safe_load(config_file)

    APP_NAME = config_data["name"]
    APP_VERSION = config_data["version"]

    title = f"{APP_NAME.upper()} v.{APP_VERSION}"

    LOGGER.info(len(title) * "_")
    LOGGER.info(title)
    LOGGER.info(len(title) * "_")

    return config_data


def log_dirs():
    LOGGER.info(f"Current directory: {CURRENT_DIR} ")


def get_machine_name():
    global MACHINE_NAME

    MACHINE_NAME = socket.gethostname()

    LOGGER.info(f"This machine's name: {MACHINE_NAME}")

    return True


def check_tasks(config_data):
    tasks = []

    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()

    folders = [scheduler.GetFolder("\\")]

    min_ts = datetime.now() - timedelta(seconds=int(config_data["run_delta"]))
    LOGGER.info(f"Oldest timestamp allowed for next run (UTC Time): {str(min_ts)[:19]}")

    while folders:
        folder = folders.pop(0)
        folders += list(folder.GetFolders(0))

        for task in folder.GetTasks(0):
            next_run = datetime.strptime(str(task.NextRunTime)[:19], "%Y-%m-%d %H:%M:%S")

            if TASK_STATE[task.State].upper() == "READY" and next_run.year != 1899:
                LOGGER.info(f"Investigating '{task.Path[1:]}'")

                if next_run < min_ts:
                    tasks.append({"path": task.Path[1:], "last_run": task.LastRunTime, "next_run": task.NextRunTime})
                    LOGGER.warning(f"'{task.Path[1:]}' has a RUN TIME IN THE PAST ! ({str(task.NextRunTime)[:19]})")
                else:
                    LOGGER.info(f"'{task.Path[1:]}' OK ({str(task.NextRunTime)[:19]})")

    return len(tasks) == 0, tasks


def notify(config_data, tasks):
    try:
        with smtplib.SMTP(host=config_data["mail"]["server"], port=int(config_data["mail"]["port"])) as server:
            msg = MIMEMultipart()

            msg["subject"] = Header(config_data["mail"]["subject"])
            msg["from"] = Header(config_data["mail"]["from"])
            msg["to"] = Header(config_data["mail"]["to"])

            body = f"{config_data['mail']['text']}: {MACHINE_NAME}\n\n"
            body += f"{config_data['mail']['list_text']}:\n"

            for task in tasks:
                body += f"- {task}\n"

            msg.attach(MIMEText(body, "plain"))

            server.send_message(msg=msg)

            LOGGER.info(f"Notifying {msg['to']} that one or more tasks has problems")

    except Exception as err:
        LOGGER.error(f"Could not notify people because of a mail server error: {err}")


def run(config_data):
    get_machine_name()

    check_ok, tasks = check_tasks(config_data=config_data)

    if not check_ok:
        notify(config_data=config_data, tasks=tasks)


def main():
    configure_logger()
    config_data = read_config()
    log_dirs()

    cron = croniter(config_data["cron"], datetime.now())
    next_run = cron.get_next(ret_type=datetime)

    LOGGER.info(f"Next run at: {next_run}")

    sleep(60 - datetime.now().second)

    while True:
        LOGGER.debug("Checking ...")

        now = datetime.now()

        if now > next_run:
            run(config_data=config_data)

            next_run = cron.get_next(ret_type=datetime)
            LOGGER.info(f"Next run at: {next_run}")

        sleep(60 - now.second)


if __name__ == "__main__":
    main()
