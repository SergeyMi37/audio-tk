"""
Voice Transcriber - Транскрипция голоса с микрофона в буфер обмена Windows
Требуемые библиотеки: pip install SpeechRecognition pyaudio pyperclip openai-whisper
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import speech_recognition as sr
import pyperclip
import io
import wave
import tempfile
import os
import whisper
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Проверка корректности установки PyAudio
try:
    import pyaudio
    logger.info("PyAudio успешно импортирован")
    
    # Проверка доступности устройств ввода
    audio = pyaudio.PyAudio()
    device_count = audio.get_device_count()
    logger.info(f"Найдено {device_count} аудиоустройств")
    
    # Поиск доступных устройств ввода
    input_devices = []
    for i in range(device_count):
        info = audio.get_device_info_by_index(i)
        if info['maxInputChannels'] > 0:  # Устройство поддерживает ввод
            input_devices.append((i, info['name']))
    
    logger.info(f"Найдено {len(input_devices)} устройств ввода")
    for device_id, name in input_devices:
        logger.info(f"  Устройство ввода {device_id}: {name}")
    
    audio.terminate()
    
    if len(input_devices) == 0:
        logger.warning("Не найдено устройств ввода звука")
    else:
        logger.info("PyAudio корректно установлен и настроен")
        
except ImportError:
    logger.error("PyAudio не установлен или не может быть импортирован")
    raise
except Exception as e:
    logger.error(f"Ошибка при проверке PyAudio: {e}")
    raise


class SafeMicrophone(sr.Microphone):
    """Обертка для безопасного использования микрофона"""
    def __enter__(self):
        try:
            return super().__enter__()
        except Exception as e:
            logger.error(f"Ошибка при входе в контекст микрофона: {e}")
            raise
    
    def __exit__(self, exc_type, exc_value, traceback):
        try:
            # Проверяем, что объект источника существует и имеет метод close
            if hasattr(self, 'stream') and self.stream:
                try:
                    self.stream.close()
                except Exception:
                    pass  # Игнорируем ошибки при закрытии
            return super().__exit__(exc_type, exc_value, traceback)
        except Exception as e:
            logger.error(f"Ошибка при выходе из контекста микрофона: {e}")
            # Не пробрасываем исключение, чтобы не прерывать нормальное завершение


class VoiceTranscriberApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Голосовой транскрибер")
        self.root.geometry("500x400")
        self.root.resizable(True, True)

        self.is_recording = False
        self.recognizer = sr.Recognizer()
        
        # Инициализация модели Whisper
        self.whisper_model = whisper.load_model("base")

        self.setup_ui()

    def setup_ui(self):
        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        title_label = ttk.Label(
            main_frame,
            text="Транскрипция голоса в текст",
            font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 10))

        # Статус
        self.status_var = tk.StringVar(value="Готов к записи")
        status_label = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            font=("Arial", 10)
        )
        status_label.pack(pady=5)

        # Индикатор записи
        self.indicator_canvas = tk.Canvas(
            main_frame,
            width=20,
            height=20,
            highlightthickness=0
        )
        self.indicator_canvas.pack(pady=5)
        self.indicator = self.indicator_canvas.create_oval(
            2, 2, 18, 18,
            fill="gray",
            outline="darkgray"
        )

        # Кнопки управления
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)

        self.record_btn = ttk.Button(
            button_frame,
            text="Начать запись",
            command=self.toggle_recording,
            width=20
        )
        self.record_btn.pack(side=tk.LEFT, padx=5)

        self.copy_btn = ttk.Button(
            button_frame,
            text="Копировать текст",
            command=self.copy_to_clipboard,
            width=20
        )
        self.copy_btn.pack(side=tk.LEFT, padx=5)

        # Выбор языка
        lang_frame = ttk.Frame(main_frame)
        lang_frame.pack(pady=5)

        ttk.Label(lang_frame, text="Язык:").pack(side=tk.LEFT, padx=5)
        self.language_var = tk.StringVar(value="ru-RU")
        language_combo = ttk.Combobox(
            lang_frame,
            textvariable=self.language_var,
            values=["ru-RU", "en-US", "de-DE", "fr-FR", "es-ES"],
            state="readonly",
            width=10
        )
        language_combo.pack(side=tk.LEFT, padx=5)

        # Текстовое поле для результата
        text_frame = ttk.LabelFrame(main_frame, text="Распознанный текст")
        text_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Scrollbar
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.text_area = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Arial", 11),
            yscrollcommand=scrollbar.set
        )
        self.text_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.config(command=self.text_area.yview)

        # Кнопка очистки
        clear_btn = ttk.Button(
            main_frame,
            text="Очистить",
            command=self.clear_text,
            width=15
        )
        clear_btn.pack(pady=5)

        # Автокопирование
        self.auto_copy_var = tk.BooleanVar(value=True)
        auto_copy_check = ttk.Checkbutton(
            main_frame,
            text="Автоматически копировать в буфер обмена",
            variable=self.auto_copy_var
        )
        auto_copy_check.pack(pady=5)

    def toggle_recording(self):
        if self.is_recording:
            self.stop_recording()
        else:
            self.start_recording()

    def start_recording(self):
        # Проверка возможности создания микрофона
        try:
            test_microphone = SafeMicrophone()
            del test_microphone  # Удаляем временный объект
        except Exception as e:
            logger.error(f"Ошибка инициализации микрофона: {e}")
            messagebox.showerror(
                "Ошибка",
                f"Ошибка инициализации микрофона: {e}\n\nУбедитесь, что микрофон подключен и доступен."
            )
            return
            
        self.is_recording = True
        self.record_btn.config(text="Остановить запись")
        self.indicator_canvas.itemconfig(self.indicator, fill="red")
        self.status_var.set("Запись... Говорите в микрофон")

        # Запуск в отдельном потоке
        self.record_thread = threading.Thread(target=self.record_audio, daemon=True)
        self.record_thread.start()

    def stop_recording(self):
        self.is_recording = False
        self.record_btn.config(text="Начать запись")
        self.indicator_canvas.itemconfig(self.indicator, fill="gray")
        self.status_var.set("Готов к записи")

    def record_audio(self):
        # Сначала выполним калибровку микрофона
        self.status_var.set("Калибровка микрофона...")
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
        except Exception as e:
            logger.error(f"Ошибка калибровки микрофона: {e}")
            self.root.after(
                0,
                messagebox.showerror,
                "Ошибка",
                f"Ошибка калибровки микрофона: {e}"
            )
            self.root.after(0, self.stop_recording)
            return

        # Основной цикл записи
        while self.is_recording:
            self.status_var.set("Слушаю... Говорите")
            
            try:
                # Создаем новый экземпляр микрофона для каждого цикла прослушивания
                # Это помогает избежать проблем с закрытием ресурсов
                with sr.Microphone() as source:
                    # Запись аудио с таймаутом
                    audio = self.recognizer.listen(
                        source,
                        timeout=5,
                        phrase_time_limit=10
                    )

                self.status_var.set("Распознавание...")

                # Преобразование аудио в формат WAV для Whisper
                wav_data = io.BytesIO()
                with wave.open(wav_data, 'wb') as wav_file:
                    wav_file.setnchannels(audio.sample_rate)
                    wav_file.setsampwidth(audio.sample_width)
                    wav_file.setframerate(audio.sample_rate)
                    wav_file.writeframes(audio.frame_data)
                
                # Сохранение временного файла
                with tempfile.NamedTemporaryFile(delete=False, suffix='.wav') as temp_file:
                    temp_file.write(wav_data.getvalue())
                    temp_filename = temp_file.name
                
                try:
                    # Распознавание через локальную модель Whisper
                    result = self.whisper_model.transcribe(temp_filename, language=self.get_whisper_language_code())
                    text = result['text']
                    
                    # Добавление текста в поле
                    self.root.after(0, self.append_text, text)
                    
                except Exception as e:
                    logger.error(f"Ошибка транскрибации Whisper: {e}")
                    self.root.after(
                        0,
                        lambda: self.status_var.set("Ошибка транскрибации, продолжаю...")
                    )
                    continue
                finally:
                    # Удаление временного файла
                    if os.path.exists(temp_filename):
                        os.unlink(temp_filename)

            except sr.WaitTimeoutError:
                # Таймаут - продолжаем слушать
                continue
            except sr.UnknownValueError:
                # Речь не распознана
                self.status_var.set("Речь не распознана, продолжаю...")
                continue
            except sr.RequestError as e:
                logger.error(f"Ошибка запроса к сервису распознавания: {e}")
                self.status_var.set("Ошибка сервиса распознавания, продолжаю...")
                continue
            except OSError as e:
                # Ошибка PyAudio, например, если микрофон занят или недоступен
                logger.error(f"Ошибка PyAudio: {e}")
                self.root.after(
                    0,
                    messagebox.showerror,
                    "Ошибка",
                    f"Ошибка аудиоустройства: {e}\n\nУбедитесь, что микрофон не используется другими приложениями."
                )
                break
            except Exception as e:
                logger.error(f"Неизвестная ошибка при прослушивании: {e}")
                self.root.after(
                    0,
                    messagebox.showerror,
                    "Ошибка",
                    f"Ошибка при прослушивании: {e}"
                )
                break

        # В конце останавливаем запись
        self.root.after(0, self.stop_recording)

    def append_text(self, text):
        """Добавляет распознанный текст в текстовое поле"""
        current_text = self.text_area.get("1.0", tk.END).strip()

        if current_text:
            self.text_area.insert(tk.END, " " + text)
        else:
            self.text_area.insert(tk.END, text)

        self.text_area.see(tk.END)

        # Автокопирование
        if self.auto_copy_var.get():
            self.copy_to_clipboard(show_message=False)

        self.status_var.set("Текст добавлен! Слушаю...")

    def copy_to_clipboard(self, show_message=True):
        """Копирует текст в буфер обмена"""
        text = self.text_area.get("1.0", tk.END).strip()

        if text:
            pyperclip.copy(text)
            if show_message:
                self.status_var.set("Текст скопирован в буфер обмена!")
        else:
            if show_message:
                messagebox.showwarning("Внимание", "Нет текста для копирования")

    def clear_text(self):
        """Очищает текстовое поле"""
        self.text_area.delete("1.0", tk.END)
        self.status_var.set("Текст очищен")

    def get_whisper_language_code(self):
        """Преобразует код языка из формата Google в формат Whisper"""
        google_to_whisper = {
            "ru-RU": "ru",
            "en-US": "en",
            "de-DE": "de",
            "fr-FR": "fr",
            "es-ES": "es"
        }
        return google_to_whisper.get(self.language_var.get(), "en")


def main():
    root = tk.Tk()
    app = VoiceTranscriberApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
