import pyvisa

class GPIBController:
    """GPIB通信制御クラス"""
    
    def __init__(self):
        """初期化"""
        self.rm = None
        self.instrument = None
        self.connected = False
        self.current_resource = None
        self.timeout = 5000
        
    def initialize(self):
        """VISAリソースマネージャーの初期化"""
        try:
            self.rm = pyvisa.ResourceManager()
            return True, "VISA初期化成功"
        except Exception as e:
            return False, f"VISA初期化失敗: {str(e)}"
    
    def list_resources(self):
        """接続可能なGPIB機器のリストを取得"""
        if self.rm is None:
            success, message = self.initialize()
            if not success:
                return False, message
        
        try:
            resources = self.rm.list_resources()
            if len(resources) == 0:
                return True, []
            return True, list(resources)
        except Exception as e:
            return False, f"リソース検索失敗: {str(e)}"
    
    def test_connection(self, test_commands=None, timeout=1000):
        """
        接続テストを実行
        
        Parameters:
        -----------
        test_commands : list
            テストコマンドのリスト（優先順位順）
        timeout : int
            各コマンドのタイムアウト時間（ms）
        
        Returns:
        --------
        success : bool
            テスト成功/失敗
        message : str
            テスト結果メッセージ
        """
        if not self.connected:
            return False, "機器が接続されていません"
        
        # デフォルトのテストコマンド
        if test_commands is None:
            test_commands = [
                "*IDN?",      # 標準的な識別コマンド
                "*OPC?",      # 操作完了確認
                "*STB?",      # ステータスバイト確認
            ]
        
        # 元のタイムアウトを保存
        original_timeout = self.instrument.timeout
        self.instrument.timeout = timeout
        
        test_result = None
        successful_command = None
        
        # 各コマンドを順番に試す
        for cmd in test_commands:
            try:
                response = self.instrument.query(cmd)
                test_result = response.strip()
                successful_command = cmd
                break  # 成功したらループを抜ける
            except pyvisa.errors.VisaIOError as e:
                # タイムアウトや通信エラーの場合は次のコマンドを試す
                continue
            except Exception as e:
                # その他のエラーも次のコマンドを試す
                continue
        
        # タイムアウトを元に戻す
        self.instrument.timeout = original_timeout
        
        if test_result is not None:
            return True, f"接続テスト成功: {successful_command} → {test_result}"
        else:
            return False, "接続テスト失敗: 全てのコマンドが応答なし（機器は接続されているが応答不可の可能性）"
    
    def connect(self, resource_name, timeout=5000, test_mode="auto", test_timeout=1000, test_commands=None, device_type="3458A"):
        """
        指定されたGPIB機器に接続
        
        Parameters:
        -----------
        resource_name : str
            接続するリソース名
        timeout : int
            タイムアウト時間(ms)
        test_mode : str
            接続確認モード("none", "auto", "idn")
        test_timeout : int
            テストコマンドのタイムアウト(ms)
        test_commands : list
            テストコマンドのリスト
        device_type : str
            機器タイプ("3458A", "3499B")でターミネーション設定を切り替え
        """
        try:
            # 既に接続している場合は切断
            if self.connected:
                self.disconnect()
            
            # リソースを開く
            self.instrument = self.rm.open_resource(resource_name)
            self.instrument.timeout = timeout
            self.timeout = timeout
            self.current_resource = resource_name
            
            # 機器タイプに応じたターミネーション設定
            if device_type == "3458A":
                # 3458A用の設定
                self.instrument.write_termination = '\n'
                self.instrument.read_termination = '\n'
                self.instrument.send_end = True
            elif device_type == "3499B":
                # 3499B用の設定
                self.instrument.write_termination = '\r\n'  # CR+LF
                self.instrument.read_termination = '\r\n'  # CR+LF
                self.instrument.send_end = True
            else:
                # デフォルト設定
                self.instrument.write_termination = '\n'
                self.instrument.read_termination = '\n'
                self.instrument.send_end = True
            
            self.connected = True
            
            # テストモードに応じた接続確認
            if test_mode == "none":
                return True, f"接続成功: {resource_name} (テストなし - open_resourceのみ)"
            
            elif test_mode == "auto":
                # 自動でテストコマンドを試す
                success, message = self.test_connection(test_commands, test_timeout)
                if success:
                    return True, f"接続成功: {message}"
                else:
                    return True, f"接続成功: {resource_name} ({message})"
            
            else:
                return False, f"不明なテストモード: {test_mode}"
                
        except Exception as e:
            self.connected = False
            self.current_resource = None
            return False, f"接続失敗: {str(e)}"
    
    def disconnect(self, go_to_local=True):
        """
        GPIB機器との接続を切断
        
        Parameters:
        -----------
        go_to_local : bool
            切断時にリモートモードを解除してローカル（パネル操作可能）に戻すか
        """
        try:
            if self.instrument:
                # リモートモード解除（パネル操作可能にする）
                if go_to_local:
                    try:
                        # GPIBローカルモードに戻す
                        self.instrument.control_ren(6)  # VI_GPIB_REN_ADDRESS_GTL
                    except Exception as e:
                        # ローカルモード復帰に失敗しても切断は続行
                        print(f"ローカルモード復帰警告: {str(e)}")
                
                self.instrument.close()
                self.connected = False
                self.current_resource = None
            return True, "切断成功（ローカルモードに復帰）"
        except Exception as e:
            return False, f"切断失敗: {str(e)}"
    
    def write(self, command):
        """GPIB機器にコマンドを送信"""
        if not self.connected:
            return False, "機器が接続されていません"
        
        try:
            self.instrument.write(command)
            return True, f"送信成功: {command}"
        except Exception as e:
            return False, f"送信失敗: {str(e)}"
    
    def query(self, command):
        """GPIB機器にクエリを送信して応答を取得"""
        if not self.connected:
            return False, "機器が接続されていません"
        
        try:
            response = self.instrument.query(command)
            return True, response.strip()
        except Exception as e:
            return False, f"クエリ失敗: {str(e)}"
    
    def read(self):
        """GPIB機器からデータを読み取り"""
        if not self.connected:
            return False, "機器が接続されていません"
        
        try:
            data = self.instrument.read()
            return True, data.strip()
        except Exception as e:
            return False, f"読み取り失敗: {str(e)}"
    
    def read_raw(self):
        """GPIB機器からバイナリデータを読み取り"""
        if not self.connected:
            return False, "機器が接続されていません"
        
        try:
            data = self.instrument.read_raw()
            return True, data
        except Exception as e:
            return False, f"読み取り失敗: {str(e)}"
    
    def query_binary_values(self, command, datatype='f', container=list):
        """バイナリ値を取得"""
        if not self.connected:
            return False, "機器が接続されていません"
        
        try:
            values = self.instrument.query_binary_values(
                command, 
                datatype=datatype, 
                container=container
            )
            return True, values
        except Exception as e:
            return False, f"バイナリ取得失敗: {str(e)}"
    
    def set_timeout(self, timeout):
        """タイムアウト値を設定"""
        if not self.connected:
            return False, "機器が接続されていません"
        
        try:
            self.instrument.timeout = timeout
            self.timeout = timeout
            return True, f"タイムアウトを{timeout}msに設定しました"
        except Exception as e:
            return False, f"タイムアウト設定失敗: {str(e)}"
    
    def get_info(self):
        """接続情報を取得"""
        if not self.connected:
            return {
                "connected": False,
                "resource": None,
                "timeout": self.timeout
            }
        
        return {
            "connected": True,
            "resource": self.current_resource,
            "timeout": self.timeout
        }
    
    def clear(self):
        """機器のバッファをクリア"""
        if not self.connected:
            return False, "機器が接続されていません"
        
        try:
            self.instrument.clear()
            return True, "バッファをクリアしました"
        except Exception as e:
            return False, f"クリア失敗: {str(e)}"
    
    def reset(self):
        """機器をリセット"""
        if not self.connected:
            return False, "機器が接続されていません"
        
        try:
            self.instrument.write("*RST")
            return True, "機器をリセットしました"
        except Exception as e:
            return False, f"リセット失敗: {str(e)}"