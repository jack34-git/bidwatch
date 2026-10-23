import os
import sys

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "settings")

if __name__ == "__main__":
    import django
    django.setup()

    command = sys.argv[1] if len(sys.argv) > 1 else ""

    if command == "initapp":
        from app import init_tables
        init_tables()
        print("공고 검색용 테이블을 생성했습니다.")

    elif command == "worker":
        from app import worker
        worker()

    else:
        from django.core.management import execute_from_command_line
        execute_from_command_line(sys.argv)
