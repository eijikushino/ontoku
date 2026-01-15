def validate_number(value, min_val=None, max_val=None):
    """数値の妥当性チェック"""
    try:
        num = float(value)
        if min_val is not None and num < min_val:
            return False, f"値が小さすぎます（最小値: {min_val}）"
        if max_val is not None and num > max_val:
            return False, f"値が大きすぎます（最大値: {max_val}）"
        return True, num
    except ValueError:
        return False, "数値を入力してください"

def validate_integer(value, min_val=None, max_val=None):
    """整数の妥当性チェック"""
    try:
        num = int(value)
        if min_val is not None and num < min_val:
            return False, f"値が小さすぎます（最小値: {min_val}）"
        if max_val is not None and num > max_val:
            return False, f"値が大きすぎます（最大値: {max_val}）"
        return True, num
    except ValueError:
        return False, "整数を入力してください"

def validate_gpib_address(address):
    """GPIBアドレスの妥当性チェック"""
    valid, num = validate_integer(address, 0, 30)
    if not valid:
        return False, "GPIBアドレスは0～30の整数です"
    return True, num

def validate_command(command):
    """コマンド文字列の妥当性チェック"""
    if not command or len(command.strip()) == 0:
        return False, "コマンドが空です"
    return True, command.strip()