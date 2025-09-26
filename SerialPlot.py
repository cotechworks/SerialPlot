import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import serial
import serial.tools.list_ports
import threading
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.animation as animation
import csv


class SerialPlotter:
    def __init__(self, root):
        self.root = root
        self.root.title("シリアルプロットツール")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.serial_port = None
        self.is_receiving = False
        self.receive_thread = None
        self.data = []
        self.last_value = None
        self.data_offset = 0  # データ表示の開始位置
        self.auto_scroll_var = tk.BooleanVar(value=True)  # チェックボックス用の変数

        self.footer_frame = tk.Frame(self.root)
        self.footer_frame.pack(fill=tk.BOTH, side=tk.BOTTOM, padx=5, pady=5)

        self.footer_label = tk.Label(
            self.footer_frame, text="最後の受信値: なし", anchor="w"
        )
        self.footer_label.pack(padx=5, pady=5, anchor=tk.W)

        self.create_widgets()

        self.fig, self.ax = plt.subplots()
        (self.line,) = self.ax.plot([], [], "b-")
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plot_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.ani = animation.FuncAnimation(
            self.fig, self.update_plot, interval=300, cache_frame_data=False
        )

    def create_widgets(self):
        # コントロール部品用のフレーム
        self.control_frame = tk.Frame(self.root)
        self.control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)

        self.port_label = tk.Label(self.control_frame, text="COMポート:")
        self.port_label.pack(padx=5, pady=5, anchor=tk.W)

        self.port_combo = ttk.Combobox(
            self.control_frame, values=self.get_serial_ports(), state="readonly"
        )
        self.port_combo.pack(padx=5, pady=5, fill=tk.X)

        self.connect_button = tk.Button(
            self.control_frame, text="接続", command=self.connect_serial
        )
        self.connect_button.pack(padx=5, pady=5, fill=tk.X)

        self.disconnect_button = tk.Button(
            self.control_frame,
            text="切断",
            command=self.disconnect_serial,
            state=tk.DISABLED,
        )
        self.disconnect_button.pack(padx=5, pady=5, fill=tk.X)

        self.xmax_label = tk.Label(self.control_frame, text="データ数:")
        self.xmax_label.pack(padx=5, pady=5, anchor=tk.W)

        self.xmax_entry = tk.Entry(self.control_frame)
        self.xmax_entry.insert(0, "100")
        self.xmax_entry.pack(padx=5, pady=5, fill=tk.X)

        self.ymin_label = tk.Label(self.control_frame, text="縦軸最小値:")
        self.ymin_label.pack(padx=5, pady=5, anchor=tk.W)

        self.ymin_entry = tk.Entry(self.control_frame)
        self.ymin_entry.insert(0, "0")
        self.ymin_entry.pack(padx=5, pady=5, fill=tk.X)

        # CSVエクスポートボタン
        self.export_csv_button = tk.Button(
            self.control_frame, text="CSVエクスポート", command=self.export_csv_data
        )
        self.export_csv_button.pack(padx=5, pady=5, fill=tk.X)

        # 横スクロール用スライダー
        self.scroll_label = tk.Label(self.control_frame, text="データスクロール:")
        self.scroll_label.pack(padx=5, pady=(10, 5), anchor=tk.W)

        self.scroll_scale = tk.Scale(
            self.control_frame,
            from_=0,
            to=0,
            orient=tk.HORIZONTAL,
            command=self.on_scroll_change
        )
        self.scroll_scale.pack(padx=5, pady=5, fill=tk.X)

        # 自動スクロールチェックボックス
        self.auto_scroll_checkbox = tk.Checkbutton(
            self.control_frame,
            text="自動スクロール",
            variable=self.auto_scroll_var,
            command=self.on_auto_scroll_toggle
        )
        self.auto_scroll_checkbox.pack(padx=5, pady=5, anchor=tk.W)

        # プロット表示用のフレーム
        self.plot_frame = tk.Frame(self.root)
        self.plot_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)

    def get_serial_ports(self):
        ports = serial.tools.list_ports.comports()
        return [port.device for port in ports]

    def connect_serial(self):
        port = self.port_combo.get()
        if not port:
            messagebox.showerror("エラー", "COMポートを選択してください。")
            return
        try:
            self.serial_port = serial.Serial(port, baudrate=115200, timeout=0.1)
            self.is_receiving = True
            self.connect_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.NORMAL)
            self.receive_thread = threading.Thread(
                target=self.read_serial_data, daemon=True
            )
            self.receive_thread.start()
        except Exception as e:
            messagebox.showerror("接続エラー", str(e))

    def disconnect_serial(self):
        self.is_receiving = False
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.connect_button.config(state=tk.NORMAL)
        self.disconnect_button.config(state=tk.DISABLED)

    def read_serial_data(self):
        while self.is_receiving and self.serial_port and self.serial_port.is_open:
            try:
                raw_bytes = self.serial_port.readline()
                line = raw_bytes.decode("utf-8", errors="ignore").strip()

                if line:
                    try:
                        value = float(line)
                        self.data.append(value)
                        self.last_value = value
                        print(f"Received value: {value}")  # デバッグ表示
                    except ValueError:
                        print(f"非数値データを受信: {line}")  # デバッグ表示
            except Exception as e:
                print(f"受信エラー: {e}")  # デバッグ表示

    def get_xmax(self):
        try:
            return int(self.xmax_entry.get())
        except ValueError:
            return 100

    def get_ymin(self):
        try:
            return float(self.ymin_entry.get())
        except ValueError:
            return 0.0

    def on_scroll_change(self, value):
        """スライダーの値が変更された際のコールバック"""
        self.data_offset = int(value)

    def on_auto_scroll_toggle(self):
        """自動スクロールチェックボックスが変更された際のコールバック"""
        if self.auto_scroll_var.get():
            # 自動スクロールが有効になった場合、最新データに移動
            if len(self.data) > self.get_xmax():
                max_offset = len(self.data) - self.get_xmax()
                self.data_offset = max_offset
                self.scroll_scale.set(max_offset)

    def export_csv_data(self):
        """現在表示されているデータをCSVファイルにエクスポートする"""
        if not self.data:
            messagebox.showinfo("情報", "エクスポートするデータがありません。")
            return

        # 現在表示されているデータを取得（update_plotと同じロジック）
        xmax = self.get_xmax()
        
        if len(self.data) <= xmax:
            plot_data = self.data
            start_index = 0
        else:
            start_index = self.data_offset
            end_index = start_index + xmax
            plot_data = self.data[start_index:end_index]

        # ファイル保存ダイアログを表示
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="CSVファイルの保存"
        )

        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    
                    # ヘッダーを書き込み
                    writer.writerow(['Index', 'Value'])
                    
                    # データを書き込み（実際のインデックスを使用）
                    for i, value in enumerate(plot_data):
                        actual_index = start_index + i
                        writer.writerow([actual_index, value])
                
                messagebox.showinfo("成功", f"データが正常にエクスポートされました。\n{file_path}\n表示範囲: {start_index}-{start_index + len(plot_data) - 1}")
                
            except Exception as e:
                messagebox.showerror("エラー", f"CSVファイルの保存中にエラーが発生しました:\n{str(e)}")

    def update_plot(self, frame):
        xmax = self.get_xmax()
        ymin = self.get_ymin()
        auto_scroll = self.auto_scroll_var.get()  # チェックボックスの状態を取得

        # スライダーに基づいてデータ範囲を取得
        if len(self.data) <= xmax:
            # 全データが表示範囲内の場合
            plot_data = self.data
            start_index = 0
        else:
            # データが表示範囲を超える場合
            total_data = len(self.data)
            max_offset = total_data - xmax
            
            # スライダーの範囲を更新
            current_max = self.scroll_scale.cget('to')
            if current_max != max_offset:
                self.scroll_scale.config(to=max_offset)
                
                # 自動スクロールが有効な場合、最新データを表示
                if auto_scroll:
                    self.data_offset = max_offset
                    self.scroll_scale.set(max_offset)
            
            # オフセットに基づいてデータを取得
            start_index = self.data_offset
            end_index = start_index + xmax
            plot_data = self.data[start_index:end_index]

        self.ax.clear()
        if plot_data:
            self.ax.plot(range(len(plot_data)), plot_data, "b-")
        
        self.ax.set_xlim(0, xmax)

        # 縦軸の最小値をテキストボックスから取得して固定
        self.ax.set_ylim(bottom=ymin)

        # データがある場合は最大値を自動調整
        if plot_data:
            self.ax.set_ylim(top=max(plot_data))

        self.ax.set_xlabel("Data Points")
        self.ax.set_ylabel("Value")
        
        # タイトルに現在の表示範囲を追加
        if len(self.data) > xmax:
            auto_status = " (Auto)" if auto_scroll else " (Manual)"
            title = f"Serial Data Plot{auto_status} (showing {start_index}-{start_index + len(plot_data) - 1} of {len(self.data)})"
        else:
            title = "Serial Data Plot"
        self.ax.set_title(title)
        
        self.canvas.draw()

        # 最後の受信値を表示
        if self.last_value is not None:
            self.footer_label.config(text=f"Last Received Value: {self.last_value}")
        else:
            self.footer_label.config(text="Last Received Value: None")

    def on_closing(self):
        # アニメーションを停止
        if hasattr(self, "ani") and self.ani:
            self.ani.event_source.stop()
            self.ani = None

        # シリアル通信を停止
        self.is_receiving = False

        # シリアルポートが開いていれば閉じる
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
            self.serial_port = None

        # スレッドが存在し、まだ動いていれば終了を待つ（最大1秒）
        if self.receive_thread and self.receive_thread.is_alive():
            self.receive_thread.join(timeout=1)

        # matplotlibのリソースをクリーンアップ
        plt.close(self.fig)

        # ウィンドウを破棄してアプリ終了
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    try:
        app = SerialPlotter(root)
        root.mainloop()
    finally:
        # 確実にPythonプロセスを終了
        import sys

        sys.exit(0)
