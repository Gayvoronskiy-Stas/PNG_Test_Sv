import openpyxl
import random
from typing import List, Dict
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CallbackContext, CommandHandler, CallbackQueryHandler


class QuizBot:
    def __init__(self, excel_path: str):
        self.questions = self.parse_excel(excel_path)
        self.user_data = {}

    @staticmethod
    def parse_excel(file_path: str) -> List[Dict]:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        questions = []

        for row in sheet.iter_rows(values_only=True):
            correct_answers = []
            if isinstance(row[0], str):
                raw_answers = row[0].replace(" и ", ",")
                for num in raw_answers.split(","):
                    if num.strip().isdigit():
                        correct_answers.append(int(num.strip()))

            question = {
                "correct_answers": correct_answers,
                "question": row[1],
                "options": [row[2], row[3], row[4], row[5]]
            }
            questions.append(question)

        return questions

    async def start(self, update: Update, context: CallbackContext) -> None:
        user_id = update.effective_user.id
        random_questions = random.sample(self.questions, min(10, len(self.questions)))
        self.user_data[user_id] = {
            "questions": random_questions,
            "current_question": 0,
            "score": 0,
            "selected_answers": []
        }
        await self.send_question(update, context)

    async def send_question(self, update: Update, context: CallbackContext) -> None:
        user_id = update.effective_user.id
        user_state = self.user_data[user_id]
        question_data = user_state["questions"][user_state["current_question"]]

        message_text = (
            f"Вопрос {user_state['current_question'] + 1}/{len(user_state['questions'])}\n\n"
            f"{question_data['question']}\n\n"
        )

        for i, option in enumerate(question_data['options'], start=1):
            checkbox = "✅" if i in user_state["selected_answers"] else "☐"
            message_text += f"\n{checkbox} Вариант {i}: {option}"

        keyboard = []
        for i in range(1, 5):
            checkbox = "☑️" if i in user_state["selected_answers"] else "⬜️"
            keyboard.append(
                [InlineKeyboardButton(f"{checkbox} Выбрать вариант {i}", callback_data=f"select_{i}")]
            )

        keyboard.append([InlineKeyboardButton("📥 Ответить", callback_data="submit")])

        if update.callback_query:
            await update.callback_query.edit_message_text(
                message_text,
                reply_markup=InlineKeyboardMarkup(keyboard))
        else:
            await update.effective_message.reply_text(
                message_text,
                reply_markup=InlineKeyboardMarkup(keyboard))

    async def handle_answer(self, update: Update, context: CallbackContext) -> None:
        query = update.callback_query
        await query.answer()
        user_id = update.effective_user.id
        data = query.data

        if data.startswith("select_"):
            option_num = int(data.split("_")[1])
            if option_num in self.user_data[user_id]["selected_answers"]:
                self.user_data[user_id]["selected_answers"].remove(option_num)
            else:
                self.user_data[user_id]["selected_answers"].append(option_num)

            await self.send_question(update, context)

        elif data == "submit":
            question_data = self.user_data[user_id]["questions"][self.user_data[user_id]["current_question"]]
            correct_answers = set(question_data["correct_answers"])
            user_answers = set(self.user_data[user_id]["selected_answers"])

            if user_answers == correct_answers:
                self.user_data[user_id]["score"] += 1

            self.user_data[user_id]["current_question"] += 1
            self.user_data[user_id]["selected_answers"] = []

            if self.user_data[user_id]["current_question"] < len(self.user_data[user_id]["questions"]):
                await self.send_question(update, context)
            else:
                await self.finish_quiz(update, context)

    async def finish_quiz(self, update: Update, context: CallbackContext) -> None:
        user_id = update.effective_user.id
        score = self.user_data[user_id]["score"]
        total = len(self.user_data[user_id]["questions"])
        result = "✅ Тест сдан!" if (score / total) >= 0.8 else "❌ Тест не сдан."

        text = f"Тест завершен!\nПравильных ответов: {score}/{total}\nРезультат: {result}"
        await update.effective_message.reply_text(text)
        del self.user_data[user_id]


def main() -> None:
    TOKEN = "7223808501:AAH_QtfPx-Kc8ge1-elYtwFQhXOZ7D8_VNU"
    EXCEL_PATH = "Test1.xlsx"

    app = Application.builder().token(TOKEN).build()
    bot = QuizBot(EXCEL_PATH)

    app.add_handler(CommandHandler("start", bot.start))
    app.add_handler(CallbackQueryHandler(bot.handle_answer))

    print("Бот запущен...")
    app.run_polling()


if __name__ == "__main__":
    main()