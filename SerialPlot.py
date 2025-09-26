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
            self.fig, self.update_plot, interval=500, cache_frame_data=False
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

        self.xmax_label = tk.Label(self.control_frame, text="横軸最大値（データ数）:")
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

    def export_csv_data(self):
        """現在のplot_dataをCSVファイルにエクスポートする"""
        if not self.data:
            messagebox.showinfo("情報", "エクスポートするデータがありません。")
            return

        # 現在表示されているデータを取得（update_plotと同じロジック）
        xmax = self.get_xmax()
        if len(self.data) > xmax:
            plot_data = self.data[-xmax:]
        else:
            plot_data = self.data

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
                    
                    # データを書き込み
                    for i, value in enumerate(plot_data):
                        writer.writerow([i, value])
                
                messagebox.showinfo("成功", f"データが正常にエクスポートされました。\n{file_path}")
                
            except Exception as e:
                messagebox.showerror("エラー", f"CSVファイルの保存中にエラーが発生しました:\n{str(e)}")

    def update_plot(self, frame):
        xmax = self.get_xmax()
        ymin = self.get_ymin()

        if len(self.data) > xmax:
            plot_data = self.data[-xmax:]
        else:
            plot_data = self.data

        self.ax.clear()
        self.ax.plot(range(len(plot_data)), plot_data, "b-")
        self.ax.set_xlim(0, xmax)

        # 縦軸の最小値をテキストボックスから取得して固定
        self.ax.set_ylim(bottom=ymin)

        # データがある場合は最大値を自動調整
        if plot_data:
            self.ax.set_ylim(top=max(plot_data))

        self.ax.set_xlabel("Data Points")
        self.ax.set_ylabel("Value")
        self.ax.set_title("Serial Data Plot")
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
