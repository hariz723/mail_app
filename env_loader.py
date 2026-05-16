import os

from dotenv import load_dotenv  # type: ignore


def get_env_variable(name, default=None, cast_type=str, env_file=".env"):
    load_dotenv(env_file)
    value = os.getenv(name, default)

    if value is None:
        return None

    return cast_type(value)


class SMTPConfig:
    smtp_host = get_env_variable("SMTP_HOST")
    smtp_port = get_env_variable("SMTP_PORT", 587, int)
    smtp_username = get_env_variable("SMTP_USERNAME")
    smtp_password = get_env_variable("SMTP_PASSWORD")
    smtp_from_email = get_env_variable("SMTP_FROM_EMAIL", smtp_username)

    @classmethod
    def validate(cls):
        if (
            not cls.smtp_host
            or not cls.smtp_username
            or not cls.smtp_password
            or not cls.smtp_from_email
        ):
            raise ValueError(
                "Please set SMTP_HOST, SMTP_USERNAME, SMTP_PASSWORD, and SMTP_FROM_EMAIL."
            )


settings = SMTPConfig()