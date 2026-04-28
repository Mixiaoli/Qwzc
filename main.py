import tkinter as tk
from tkinter import messagebox
import keyboard
import time
import pyperclip

from core import CollectorCore
from excel import save_excel
from config import load_config, save_config


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("企微智齿采集工具-Mixiao")
        self.root.geometry("520x360")

        self.core = CollectorCore(self.update_log)

        # 👉 状态控制
        self.last_time = 0
        self.running = False

        # 👉 读取配置
        self.config_data = load_config()

        tk.Label(root, text="受理客服邮箱：").pack()

        self.email_entry = tk.Entry(root, width=40)
        self.email_entry.pack(pady=5)

        # 👉 自动填充
        last_email = self.config_data.get("service_email", "")
        self.email_entry.insert(0, last_email)

        tk.Button(root, text="启动采集", command=self.start).pack(pady=5)
        tk.Button(root, text="导出Excel（xlsx）", command=self.export).pack(pady=5)

        self.log = tk.Text(root, height=12)
        self.log.pack(fill="both", expand=True)

    def update_log(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)

    # ==========================
    # ✅ F2 触发逻辑（完整版）
    # ==========================
    def safe_hotkey(self):
        now = time.time()

        # 👉 防连点
        if now - self.last_time < 0.8:
            return

        if self.running:
            self.update_log("⚠️ 正在处理，请稍等")
            return

        self.last_time = now
        self.running = True

        try:
            self.update_log("🔥 F2触发")

            # ==========================
            # 1️⃣ 记录旧剪贴板
            # ==========================
            old_text = pyperclip.paste()

            # ==========================
            # 2️⃣ 自动复制
            # ==========================
            keyboard.press_and_release("ctrl+c")
            time.sleep(0.25)

            # ==========================
            # 3️⃣ 获取新数据
            # ==========================
            new_text = pyperclip.paste().strip()

            # 👉 显示读取到的数据（关键！！）
            self.update_log(f"📋 读取到内容: {new_text}")

            # ==========================
            # 4️⃣ 判断是否复制成功
            # ==========================
            if not new_text or new_text == old_text:
                self.update_log("⚠️ 未检测到新复制内容，请确认已选中")
                return

            # ==========================
            # 5️⃣ 传给核心逻辑
            # ==========================
            self.core.on_hotkey(new_text)

        except Exception as e:
            self.update_log(f"❌ 错误: {e}")

        finally:
            self.running = False

    def start(self):
        email = self.email_entry.get().strip()

        # 👉 保存邮箱
        self.config_data["service_email"] = email
        save_config(self.config_data)

        self.core.set_email(email)

        # 👉 清除旧热键
        keyboard.unhook_all()

        # 👉 注册 F2
        keyboard.add_hotkey("f2", self.safe_hotkey)

        self.update_log("🚀 已启动")
        self.update_log("👉 F2：选中 → 自动复制 → 自动识别")

    def export(self):
        msg = save_excel(self.core.data)
        messagebox.showinfo("提示", msg)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()