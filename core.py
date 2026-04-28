class CollectorCore:
    def __init__(self, update_ui_callback):
        self.current_name = None
        self.data = []
        self.update_ui = update_ui_callback
        self.service_email = ""
        self.last_text = ""

    def set_email(self, email):
        self.service_email = email

    # ==========================
    # ✅ 接收 main 传入的 text
    # ==========================
    def on_hotkey(self, text):
        text = text.strip()

        # 👉 防空 & 防重复
        if not text or text == self.last_text:
            return

        self.last_text = text

        # ==========================
        # 第一步：用户名
        # ==========================
        if self.current_name is None:
            self.current_name = text
            self.update_ui(f"👤 用户: {text}")
            return

        # ==========================
        # 第二步：聊天内容
        # ==========================
        name = self.current_name
        msg = text

        self.update_ui(f"💬 内容: {msg}")

        ticket = {
            "工单标题（必填）": msg[:30],
            "工单描述（必填）": msg,
            "客户昵称（昵称、邮箱或手机至少填写一个）": name,
            "客户邮箱": name if "@" in name else "",

            # ==========================
            # ✅ 固定字段（你模板要求）
            # ==========================
            "客户手机": "",
            "状态（必填）": "已解决",
            "优先级（必填）": "低",
            "受理技能组": "[厦门&二级校]IT支持组",
            "受理客服": self.service_email,
            "反馈渠道": "企业微信",
            "请求类型（400）": "服务请求",
            "解决方案": "",
            "挂单状态": "关闭",
            "受理角色": "分校IT解决",
            "三线处理": "",
            "网络组意见": "",
        }

        self.data.append(ticket)

        self.update_ui(f"📥 已记录 {len(self.data)} 条")

        # 👉 重置（准备下一条）
        self.current_name = None