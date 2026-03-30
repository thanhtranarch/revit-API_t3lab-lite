# script.py
# -*- coding: utf-8 -*-
from pyrevit import forms
import openai

# ✅ Đặt API Key của bạn tại đây
openai.api_key = "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"  # ← THAY BẰNG KEY CỦA BẠN

# 👉 Hộp thoại hỏi người dùng nhập prompt
user_prompt = forms.ask_for_string(
    default="What can GPT-4o do in Revit?",
    prompt="Enter your question for ChatGPT-4o:"
)

if user_prompt:
    try:
        # 🚀 Gửi câu hỏi tới GPT-4o
        response = openai.ChatCompletion.create(
            model="gpt-4o",  # ← Dùng mô hình mới nhất
            messages=[
                {"role": "system", "content": "You are a helpful assistant that works inside Autodesk Revit."},
                {"role": "user", "content": user_prompt}
            ]
        )
        reply = response.choices[0].message.content.strip()

        # 📤 Hiển thị phản hồi
        forms.alert(reply, title="ChatGPT-4o Response", warn_icon=False)

    except Exception as e:
        forms.alert("❌ Error: {}", title="ChatGPT Error".format(e), warn_icon=True)
