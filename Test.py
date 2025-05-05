import openpyxl
import random
import sqlite3
import os
import logging
import argparse
import json
from typing import List, Dict
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CallbackContext, CommandHandler, CallbackQueryHandler
from dotenv import load_dotenv

# Configure logging
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)


class QuizBot:
    def __init__(self, excel_path: str):
        """Initialize the bot with Excel file."""
        self.questions = self.parse_excel(excel_path)
        self.conn = sqlite3.connect("quiz.db")
        self.create_table()
        logger.info(f"Bot initialized with {len(self.questions)} questions")

    def create_table(self):
        """Create SQLite table for user data and ensure user_answers column exists."""
        try:
            # Create table if it doesn't exist
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS user_data (
                    user_id INTEGER PRIMARY KEY,
                    questions TEXT,
                    current_question INTEGER,
                    score INTEGER,
                    selected_answers TEXT
                )
            """)

            # Check if user_answers column exists
            cursor = self.conn.cursor()
            cursor.execute("PRAGMA table_info(user_data)")
            columns = [col[1] for col in cursor.fetchall()]

            if "user_answers" not in columns:
                logger.info("Adding user_answers column to user_data table")
                cursor.execute("ALTER TABLE user_data ADD COLUMN user_answers TEXT")
                self.conn.commit()
                logger.info("user_answers column added successfully")

        except sqlite3.Error as e:
            logger.error(f"Failed to create or update SQLite table: {e}")
            raise

    @staticmethod
    def parse_excel(file_path: str) -> List[Dict]:
        """Parse questions from Excel file."""
        try:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active
            questions = []
            skipped_rows = []

            for row_num, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                # Skip if answer or question is empty
                if not row[0] or not row[1]:
                    skipped_rows.append((row_num, "Empty answer or question"))
                    continue

                # Normalize answer options, replace None with "Нет ответа"
                options = [str(opt) if opt is not None else "Нет ответа" for opt in row[2:6]]
                if not all(opt.strip() for opt in options):
                    skipped_rows.append((row_num, "One or more answer options empty after normalization"))
                    continue

                correct_answers = []
                raw_answer = str(row[0]).strip().replace("  ", " ")  # Normalize spaces

                # Handle single correct answer (e.g., "1", "2", "3", "4")
                if raw_answer in ["1", "2", "3", "4"]:
                    correct_answers.append(int(raw_answer))
                # Handle dual correct answers (e.g., "1 и 3")
                elif " и " in raw_answer:
                    nums = raw_answer.replace(" и ", ",").split(",")
                    for num in nums:
                        num = num.strip()
                        if num in ["1", "2", "3", "4"]:
                            correct_answers.append(int(num))
                        else:
                            skipped_rows.append((row_num, f"Invalid number in answers: {num}"))
                            break
                else:
                    skipped_rows.append((row_num, f"Invalid answer format: {raw_answer}"))
                    continue

                if not correct_answers:
                    skipped_rows.append((row_num, "No valid correct answers parsed"))
                    continue

                question_text = str(row[1]).strip()
                question = {
                    "correct_answers": correct_answers,
                    "question": question_text,
                    "options": options
                }
                questions.append(question)
                logger.debug(f"Parsed question {row_num}: {question_text[:50]}...")

            if skipped_rows:
                logger.warning(f"Skipped {len(skipped_rows)} rows: {skipped_rows}")
            if len(questions) <= 10:
                raise ValueError(
                    f"Excel file contains only {len(questions)} valid questions. "
                    "Must contain more than 10 questions. "
                    f"Skipped rows: {skipped_rows}"
                )
            if len(questions) > 500:
                raise ValueError("Excel file contains too many questions (max 500)")
            logger.info(f"Parsed {len(questions)} questions from {file_path}")
            return questions
        except FileNotFoundError:
            logger.error(f"Excel file {file_path} not found")
            raise FileNotFoundError(f"File {file_path} not found")
        except Exception as e:
            logger.error(f"Error reading Excel: {e}")
            raise Exception(f"Error reading Excel: {e}")

    def get_user_data(self, user_id: int) -> Dict:
        """Retrieve user data from SQLite."""
        try:
            cursor = self.conn.cursor()
            cursor.execute("SELECT * FROM user_data WHERE user_id = ?", (user_id,))
            data = cursor.fetchone()
            if data:
                return {
                    "questions": json.loads(data[1]),
                    "current_question": data[2],
                    "score": data[3],
                    "selected_answers": json.loads(data[4]),
                    "user_answers": json.loads(data[5]) if data[5] else []
                }
            return None
        except sqlite3.Error as e:
            logger.error(f"Failed to get user data: {e}")
            raise

    def save_user_data(self, user_id: int, user_data: Dict):
        """Save user data to SQLite."""
        try:
            cursor = self.conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO user_data (user_id, questions, current_question, score, selected_answers, user_answers)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                user_id,
                json.dumps(user_data["questions"]),
                user_data["current_question"],
                user_data["score"],
                json.dumps(user_data["selected_answers"]),
                json.dumps(user_data.get("user_answers", []))
            ))
            self.conn.commit()
        except sqlite3.Error as e:
            logger.error(f"Failed to save user data: {e}")
            raise

    def delete_user_data(self, user_id: int):
        """Delete user data from SQLite."""
        try:
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM user_data WHERE user_id = ?", (user_id,))
            self.conn.commit()
        except sqlite3.Error as e:
            logger.error(f"Failed to delete user data: {e}")
            raise

    async def start(self, update: Update, context: CallbackContext) -> None:
        """Handle /start command."""
        try:
            user_id = update.effective_user.id
            random_questions = random.sample(self.questions, 10)
            user_data = {
                "questions": random_questions,
                "current_question": 0,
                "score": 0,
                "selected_answers": [],
                "user_answers": []
            }
            self.save_user_data(user_id, user_data)

            welcome_text = (
                "Добро пожаловать в PNG-test bot! 📝\n"
                "Вы получите 10 вопросов с 4 вариантами ответов. "
                "Выберите один или несколько вариантов, затем нажмите '📥 Ответить'.\n"
                "Для успешного прохождения необходимо правильно ответить на >=80% вопросов.\n"
                "В конце можно ознакомиться с результатами и разобрать ошибки.\n"
                "*При возникновении технических проблем с ботом обращаться к разработчику @Gayvoronskiy_Stas"
            )
            await update.effective_message.reply_text(welcome_text)
            await self.send_question(update, context)
            logger.info(f"User {user_id} started quiz with 10 questions")
        except Exception as e:
            logger.error(f"Error in start: {e}")
            await update.effective_message.reply_text("Произошла ошибка. Попробуйте снова.")

    def format_question(self, user_state: Dict, question_data: Dict) -> str:
        """Format question text with options, displaying full text and separators."""
        message_text = (
            f"Вопрос {user_state['current_question'] + 1}/10\n\n"
            f"{question_data['question']}\n\n"
        )
        for i, option in enumerate(question_data['options'], start=1):
            checkbox = "✅" if i in user_state["selected_answers"] else "⬜️"
            message_text += f"{checkbox} Вариант {i}: {option}\n"
            if i < len(question_data['options']):
                message_text += f"{'-' * 50}\n"
        return message_text

    async def send_question(self, update: Update, context: CallbackContext) -> None:
        """Send current question to user, splitting if necessary."""
        try:
            user_id = update.effective_user.id
            user_state = self.get_user_data(user_id)
            if not user_state:
                await update.effective_message.reply_text("Сессия истекла. Нажмите /start.")
                return

            question_data = user_state["questions"][user_state["current_question"]]
            message_text = self.format_question(user_state, question_data)

            keyboard = []
            for i in range(1, 5):
                checkbox = "☑️" if i in user_state["selected_answers"] else "⬜️"
                keyboard.append(
                    [InlineKeyboardButton(f"{checkbox} Вариант {i}", callback_data=f"select_{i}")]
                )
            keyboard.append([InlineKeyboardButton("📥 Ответить", callback_data="submit")])

            # Split message if it exceeds Telegram's 4096 character limit
            if len(message_text) > 4096:
                parts = []
                current_part = ""
                lines = message_text.split("\n")
                for line in lines:
                    if len(current_part) + len(line) + 1 > 4096:
                        parts.append(current_part)
                        current_part = line + "\n"
                    else:
                        current_part += line + "\n"
                if current_part:
                    parts.append(current_part)

                for i, part in enumerate(parts):
                    if update.callback_query and i == 0:
                        await update.callback_query.edit_message_text(
                            part, reply_markup=InlineKeyboardMarkup(keyboard))
                    else:
                        await update.effective_message.reply_text(
                            part, reply_markup=InlineKeyboardMarkup(keyboard) if i == 0 else None)
            else:
                if update.callback_query:
                    await update.callback_query.edit_message_text(
                        message_text, reply_markup=InlineKeyboardMarkup(keyboard))
                else:
                    await update.effective_message.reply_text(
                        message_text, reply_markup=InlineKeyboardMarkup(keyboard))
        except Exception as e:
            logger.error(f"Error in send_question: {e}")
            await update.effective_message.reply_text("Произошла ошибка. Попробуйте снова.")

    async def handle_answer(self, update: Update, context: CallbackContext) -> None:
        """Handle user answer or selection."""
        try:
            query = update.callback_query
            await query.answer()
            user_id = update.effective_user.id
            user_state = self.get_user_data(user_id)
            if not user_state:
                await query.message.reply_text("Сессия истекла. Нажмите /start.")
                return

            data = query.data
            if data == "restart":
                await self.start(update, context)
                return
            elif data.startswith("select_"):
                option_num = int(data.split("_")[1])
                if option_num in user_state["selected_answers"]:
                    user_state["selected_answers"].remove(option_num)
                else:
                    user_state["selected_answers"].append(option_num)
                self.save_user_data(user_id, user_state)
                await self.send_question(update, context)
            elif data == "submit":
                question_data = user_state["questions"][user_state["current_question"]]
                correct_answers = set(question_data["correct_answers"])
                user_answers = set(user_state["selected_answers"])

                if "user_answers" not in user_state:
                    user_state["user_answers"] = []
                user_state["user_answers"].append(list(user_answers))

                if user_answers == correct_answers:
                    user_state["score"] += 1

                user_state["current_question"] += 1
                user_state["selected_answers"] = []
                self.save_user_data(user_id, user_state)

                if user_state["current_question"] < len(user_state["questions"]):
                    await self.send_question(update, context)
                else:
                    await self.finish_quiz(update, context)
        except Exception as e:
            logger.error(f"Error in handle_answer: {e}")
            await query.message.reply_text("Произошла ошибка. Попробуйте снова.")

    async def finish_quiz(self, update: Update, context: CallbackContext) -> None:
        """Finish quiz and show results, listing only incorrect questions with enhanced formatting."""
        try:
            user_id = update.effective_user.id
            user_state = self.get_user_data(user_id)
            if not user_state:
                await update.effective_message.reply_text("Сессия истекла. Нажмите /start.")
                return

            score = user_state["score"]
            total = len(user_state["questions"])
            result = "✅ Тест сдан!" if (score / total) >= 0.8 else "❌ Тест не сдан."

            text = f"Тест завершен!\nПравильных ответов: {score}/{total}\nРезультат: {result}\n\n"

            incorrect_questions = []
            for i, (question_data, user_answers) in enumerate(zip(user_state["questions"], user_state["user_answers"]),
                                                              1):
                correct_answers = set(question_data["correct_answers"])
                user_answers = set(user_answers)
                if user_answers != correct_answers:
                    correct_answer_texts = [question_data["options"][ans - 1] for ans in correct_answers]
                    incorrect_questions.append(
                        f"❓ Вопрос {i}: {question_data['question']}\n"
                        f"✅ Правильные ответы: {', '.join(correct_answer_texts)}\n"
                        f"{'-' * 50}\n"
                    )

            if incorrect_questions:
                text += "Неправильные ответы:\n\n" + "".join(incorrect_questions)
            else:
                text += "Все ответы правильные! 🎉\n"

            # Split results if they exceed Telegram's 4096 character limit
            if len(text) > 4096:
                parts = []
                current_part = ""
                lines = text.split("\n")
                for line in lines:
                    if len(current_part) + len(line) + 1 > 4096:
                        parts.append(current_part)
                        current_part = line + "\n"
                    else:
                        current_part += line + "\n"
                if current_part:
                    parts.append(current_part)

                for i, part in enumerate(parts):
                    await update.effective_message.reply_text(
                        part,
                        reply_markup=InlineKeyboardMarkup(
                            [[InlineKeyboardButton("🔄 Начать заново", callback_data="restart")]]) if i == len(
                            parts) - 1 else None
                    )
            else:
                await update.effective_message.reply_text(
                    text,
                    reply_markup=InlineKeyboardMarkup(
                        [[InlineKeyboardButton("🔄 Начать заново", callback_data="restart")]])
                )

            self.delete_user_data(user_id)
            logger.info(f"User {user_id} finished quiz with score {score}/{total}")
        except Exception as e:
            logger.error(f"Error in finish_quiz: {e}")
            await update.effective_message.reply_text("Произошла ошибка. Попробуйте снова.")


def main():
    """Main function to run the bot."""
    try:
        parser = argparse.ArgumentParser(description="Run Telegram quiz bot")
        parser.add_argument("--excel", default="Test1.xlsx", help="Path to Excel file")
        args = parser.parse_args()

        load_dotenv()
        TOKEN = os.getenv("TELEGRAM_TOKEN")
        if not TOKEN:
            logger.error("Telegram token not provided")
            raise ValueError("Telegram token not provided")

        app = Application.builder().token(TOKEN).build()
        bot = QuizBot(args.excel)

        app.add_handler(CommandHandler("start", bot.start))
        app.add_handler(CallbackQueryHandler(bot.handle_answer))

        logger.info("Bot started")
        app.run_polling()
    except Exception as e:
        logger.error(f"Failed to start bot: {e}")
        raise


if __name__ == "__main__":
    main()
