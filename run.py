import os
import subprocess
import sys

ALLOWED_FREQUENCIES = ['DAILY', 'WEEKLY', 'MONTHLY']
ALLOWED_MODES = ['ALL', 'PDF', 'ZALO']


def get_env(name, default=''):
    return os.environ.get(name, default).strip()


def validate_frequency(value):
    frequency = value.strip().upper() or 'DAILY'
    if frequency not in ALLOWED_FREQUENCIES:
        raise SystemExit(f"FREQUENCY không hợp lệ: {value}. Chọn trong {ALLOWED_FREQUENCIES}")
    return frequency


def validate_mode(value):
    mode = value.strip().upper() or 'ALL'
    if mode not in ALLOWED_MODES:
        raise SystemExit(f"JOB_MODE không hợp lệ: {value}. Chọn trong {ALLOWED_MODES}")
    return mode


def run_script(script_path, args=None):
    args = args or []
    cmd = [sys.executable, script_path] + args
    print(f"🔧 Running: {' '.join(cmd)}")
    return_code = subprocess.call(cmd)
    if return_code != 0:
        raise SystemExit(f"Script {script_path} failed with exit code {return_code}")
    return return_code


def run_script_safe(script_path, args=None):
    """
    Chạy script nhưng không thất bại toàn bộ workflow nếu script này fails.
    Dùng cho các job non-critical như summarize log.
    """
    args = args or []
    cmd = [sys.executable, script_path] + args
    print(f"🔧 Running (safe mode): {' '.join(cmd)}")
    try:
        return_code = subprocess.call(cmd)
        if return_code != 0:
            print(f"⚠️  Script {script_path} returned exit code {return_code}, but continuing workflow...")
        return return_code
    except Exception as e:
        print(f"⚠️  Error running {script_path}: {e}, but continuing workflow...")
        return 1


if __name__ == '__main__':
    frequency = validate_frequency(get_env('FREQUENCY', 'DAILY'))
    job_mode = validate_mode(get_env('JOB_MODE', 'ALL'))

    print(f"🚀 START workflow with JOB_MODE={job_mode}, FREQUENCY={frequency}")

    if job_mode in ('ALL', 'PDF'):
        run_script('main.py')

    if job_mode in ('ALL', 'ZALO'):
        run_script('zalo_sender.py', ['--frequency', frequency])

    # Luôn chạy summarize log ở cuối, dù PDF/Zalo có thất bại hay không
    run_script_safe('job_summarize_log.py')

    print('✅ Workflow finished successfully')
