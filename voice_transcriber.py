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
import struct
import whisper
import logging
import subprocess

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', encoding='utf-8')
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
        logger.info(f"  Устройство ввода {device_id}: {name.encode('cp1251').decode('utf-8')}")
    
    audio.terminate()
    
    if len(input_devices) == 0:
        logger.warning("Не найдено устройств ввода звука")
    else:
        logger.info("PyAudio корректно установлен и настроен")
    
    # Проверка наличия ffmpeg
    try:
        subprocess.run(['ffmpeg', '-version'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        logger.info("ffmpeg успешно найден")
    except (subprocess.SubprocessError, FileNotFoundError):
        logger.warning("ffmpeg не найден. Whisper может не работать корректно. Установите ffmpeg для корректной работы программы.")
        
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
        self.root.title("Звукозапись")
        self.root.geometry("500x400")
        self.root.resizable(True, True)

        self.is_recording = False
        self.recognizer = sr.Recognizer()
        # Устанавливаем начальный порог чувствительности
        self.recognizer.energy_threshold = 400  # Можно настроить под конкретную среду
        
        # Получение списка устройств ввода
        self.input_devices = self.get_input_devices()
        self.selected_device_id = None  # По умолчанию используем устройство по умолчанию
        
        # Инициализация модели Whisper
        self.available_models = ["tiny", "base", "small", "medium", "large"]
        self.whisper_model = whisper.load_model("base")  # Модель по умолчанию
        
        # Буфер для аудио данных
        self.audio_buffer = []

        self.setup_ui()

    def setup_ui(self):
        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Настройка сетки
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(11, weight=1)  # Для текстового поля

        # Заголовок (занимает всю ширину)
        title_label = ttk.Label(
            main_frame,
            text="Транскрипция голоса в текст",
            font=("Arial", 1, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))

        # Статус и индикатор (в одной строке)
        self.status_var = tk.StringVar(value="Готов к записи")
        status_label = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            font=("Arial", 10)
        )
        status_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=2)

        # Индикатор записи
        self.indicator_canvas = tk.Canvas(
            main_frame,
            width=20,
            height=20,
            highlightthickness=0
        )
        self.indicator_canvas.grid(row=1, column=2, sticky="e", pady=2)
        self.indicator = self.indicator_canvas.create_oval(
            2, 2, 18, 18,
            fill="gray",
            outline="darkgray"
        )

        # Кнопки управления (в одной строке)
        self.record_btn = ttk.Button(
            main_frame,
            text="Начать запись",
            command=self.toggle_recording,
            width=20
        )
        self.record_btn.grid(row=2, column=0, padx=2, pady=5, sticky="ew")

        self.copy_btn = ttk.Button(
            main_frame,
            text="Копировать текст",
            command=self.copy_to_clipboard,
            width=20
        )
        self.copy_btn.grid(row=2, column=1, padx=2, pady=5, sticky="ew")

        # Кнопка очистки
        clear_btn = ttk.Button(
            main_frame,
            text="Очистить",
            command=self.clear_text,
            width=15
        )
        clear_btn.grid(row=2, column=2, padx=2, pady=5, sticky="ew")

        # Язык (в одной строке)
        ttk.Label(main_frame, text="Язык:").grid(row=3, column=0, sticky="w", padx=2, pady=2)
        self.language_var = tk.StringVar(value="ru-RU")
        language_combo = ttk.Combobox(
            main_frame,
            textvariable=self.language_var,
            values=["ru-RU", "en-US", "de-DE", "fr-FR", "es-ES"],
            state="readonly",
            width=10
        )
        language_combo.grid(row=3, column=1, sticky="w", padx=2, pady=2)

        # Автокопирование (в той же строке, что и язык)
        self.auto_copy_var = tk.BooleanVar(value=True)
        auto_copy_check = ttk.Checkbutton(
            main_frame,
            text="Автокопирование",
            variable=self.auto_copy_var
        )
        auto_copy_check.grid(row=3, column=2, sticky="w", padx=2, pady=2)

        # Устройство ввода (в одной строке)
        ttk.Label(main_frame, text="Устройство ввода:").grid(row=4, column=0, sticky="w", padx=2, pady=2)
        
        # Подготовка значений для комбобокса (форматируем названия устройств)
        device_values = [f"{device_id}: {name.encode('cp1251').decode('utf-8')}" for device_id, name in self.input_devices]
        if not device_values:
            device_values = ["Нет доступных устройств"]
        
        self.device_var = tk.StringVar()
        if self.input_devices:
            # Устанавливаем первое устройство как выбранное по умолчанию
            self.device_var.set(device_values[0])
        else:
            self.device_var.set("Нет доступных устройств")
            
        self.device_combo = ttk.Combobox(
            main_frame,
            textvariable=self.device_var,
            values=device_values,
            state="readonly",
            width=30
        )
        self.device_combo.grid(row=4, column=1, columnspan=2, sticky="ew", padx=2, pady=2)
        self.device_combo.bind("<<ComboboxSelected>>", self.on_device_selected)

        # Чувствительность микрофона (в одной строке)
        ttk.Label(main_frame, text="Чувствительность:").grid(row=5, column=0, sticky="w", padx=2, pady=2)
        
        self.sensitivity_var = tk.DoubleVar(value=400.0)
        self.sensitivity_scale = ttk.Scale(
            main_frame,
            from_=50,
            to_=1000,
            orient=tk.HORIZONTAL,
            variable=self.sensitivity_var,
            length=200
        )
        self.sensitivity_scale.grid(row=5, column=1, sticky="ew", padx=2, pady=2)
        
        self.sensitivity_label = ttk.Label(main_frame, text="400")
        self.sensitivity_label.grid(row=5, column=2, sticky="w", padx=2, pady=2)
        
        # Обновление метки при изменении значения
        self.sensitivity_var.trace_add('write', self.update_sensitivity_label)

        # Модель распознавания (в одной строке)
        ttk.Label(main_frame, text="Модель распознавания:").grid(row=6, column=0, sticky="w", padx=2, pady=2)
        
        self.model_var = tk.StringVar(value="base")
        model_combo = ttk.Combobox(
            main_frame,
            textvariable=self.model_var,
            values=self.available_models,
            state="readonly",
            width=10
        )
        model_combo.grid(row=6, column=1, sticky="w", padx=2, pady=2)
        model_combo.bind("<<ComboboxSelected>>", self.on_model_selected)

        # Пустая строка для разделения
        ttk.Separator(main_frame, orient="horizontal").grid(row=7, column=0, columnspan=3, sticky="ew", pady=5)

        # Текстовое поле для результата (занимает оставшееся пространство)
        text_frame = ttk.LabelFrame(main_frame, text="Распознанный текст")
        text_frame.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=5)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)

        # Scrollbar
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.text_area = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Arial", 11)
        )
        self.text_area.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        scrollbar.config(command=self.text_area.yview)

    def toggle_recording(self):
        if self.is_recording:
            self.finish_recording()
        else:
            self.start_recording()

    def start_recording(self):
        # Проверка возможности создания микрофона с выбранным устройством
        try:
            test_microphone = SafeMicrophone(device_index=self.selected_device_id)
            del test_microphone  # Удаляем временный объект
        except Exception as e:
            logger.error(f"Ошибка инициализации микрофона: {e}")
            messagebox.showerror(
                "Ошибка",
                f"Ошибка инициализации микрофона: {e}\n\nУбедитесь, что микрофон подключен и доступен."
            )
            return
            
        self.is_recording = True
        self.record_btn.config(text="Закончить запись")
        self.indicator_canvas.itemconfig(self.indicator, fill="red")
        self.status_var.set("Запись... Говорите в микрофон")

        # Очистка аудио буфера перед началом новой записи
        self.audio_buffer = []

        # Запуск в отдельном потоке
        self.record_thread = threading.Thread(target=self.record_audio, daemon=True)
        self.record_thread.start()

    def stop_recording(self):
        self.is_recording = False
        self.record_btn.config(text="Начать запись")
        self.indicator_canvas.itemconfig(self.indicator, fill="gray")
        self.status_var.set("Готов к записи")
        
    def finish_recording(self):
        self.is_recording = False
        self.record_btn.config(text="Обработка...")
        self.indicator_canvas.itemconfig(self.indicator, fill="yellow")
        self.status_var.set("Обработка записи...")
        
        # Запуск транскрибации в отдельном потоке
        transcribe_thread = threading.Thread(target=self.transcribe_audio_buffer, daemon=True)
        transcribe_thread.start()

    def record_audio(self):
        # Сначала выполним калибровку микрофона
        self.status_var.set("Калибровка микрофона...")
        try:
            with sr.Microphone(device_index=self.selected_device_id) as source:
                # Обновляем порог чувствительности до калибровки
                self.update_energy_threshold()
                self.recognizer.adjust_for_ambient_noise(source, duration=1.0)  # Увеличиваем время калибровки
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
                with sr.Microphone(device_index=self.selected_device_id) as source:
                    # Запись аудио с таймаутом
                    audio = self.recognizer.listen(
                        source,
                        timeout=3,           # Уменьшаем таймаут
                        phrase_time_limit=15 # Увеличиваем максимальное время фразы
                    )

                self.status_var.set("Распознавание...")
                
                # Логирование параметров аудио
                logger.info(f"Параметры аудио: sample_rate={audio.sample_rate}, sample_width={audio.sample_width}, frame_data size={len(audio.frame_data)}")

                # Преобразование аудио в формат WAV для Whisper
                wav_data = io.BytesIO()
                
                # Проверяем корректность параметров аудио для WAV формата
                if audio.sample_width not in [1, 2, 4]:
                    logger.warning(f"Неподдерживаемый sample_width: {audio.sample_width}. Пропускаем фрейм.")
                    continue
                
                if audio.sample_rate <= 0 or audio.sample_rate > 192000:  # Максимальная частота дискретизации для WAV
                    logger.warning(f"Неподдерживаемая частота дискретизации: {audio.sample_rate}. Пропускаем фрейм.")
                    continue
                
                # speech_recognition обычно возвращает моно-аудио, поэтому количество каналов = 1
                # Проверка количества каналов не требуется, так как оно фиксировано
                
                try:
                    with wave.open(wav_data, 'wb') as wav_file:
                        # Устанавливаем параметры аудио
                        wav_file.setnchannels(1)  # Обычно микрофон - моно (1 канал)
                        wav_file.setsampwidth(audio.sample_width)
                        wav_file.setframerate(audio.sample_rate)
                        wav_file.writeframes(audio.frame_data)
                except struct.error as e:
                    logger.error(f"Ошибка при записи аудиофрейма: {e}", exc_info=True)
                    # Пропускаем этот фрейм и переходим к следующей итерации основного цикла
                    continue
                
                # Сохранение временного файла
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.wav') as temp_file:
                        temp_file.write(wav_data.getvalue())
                        temp_filename = temp_file.name
                    logger.info(f"Временный файл создан: {temp_filename}")
                except Exception as e:
                    logger.error(f"Ошибка при создании временного файла: {e}", exc_info=True)
                    continue
                
                try:
                    # Проверяем существование временного файла перед транскрибацией
                    if not os.path.exists(temp_filename):
                        logger.error(f"Временный файл {temp_filename} не существует")
                        continue
                    
                    # Распознавание через локальную модель Whisper
                    logger.info(f"Начинается транскрибация файла {temp_filename} с языком {self.get_whisper_language_code()}")
                    result = self.whisper_model.transcribe(temp_filename, language=self.get_whisper_language_code())
                    text = result['text']
                    logger.info(f"Транскрибация завершена. Результат: {text}")
                    
                    # Добавление текста в поле
                    self.root.after(0, self.append_text, text)
                    
                except FileNotFoundError as e:
                    # Ошибка может быть вызвана отсутствием ffmpeg или другого необходимого компонента
                    logger.error("Ошибка: необходимый компонент не найден. Установите ffmpeg для корректной работы Whisper.", exc_info=True)
                    self.root.after(
                        0,
                        lambda: self.status_var.set("Ошибка: отсутствует необходимый компонент (ffmpeg), см. логи...")
                    )
                    continue
                except Exception as e:
                    # Проверяем, содержит ли ошибка указание на отсутствие файла, что может указывать на отсутствие ffmpeg
                    if "file" in str(e).lower() and "not found" in str(e).lower() or isinstance(e, FileNotFoundError):
                        logger.error("Ошибка: необходимый компонент не найден. Установите ffmpeg для корректной работы Whisper.", exc_info=True)
                        self.root.after(
                            0,
                            lambda: self.status_var.set("Ошибка: отсутствует необходимый компонент (ffmpeg), см. логи...")
                        )
                        continue
                    logger.error(f"Ошибка транскрибации Whisper: {e}", exc_info=True)
                    self.root.after(
                        0,
                        lambda: self.status_var.set("Ошибка транскрибации, продолжаю...")
                    )
                    continue
                except Exception as e:
                    logger.error(f"Ошибка транскрибации Whisper: {e}", exc_info=True)  # Добавляем информацию об исключении
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

    def transcribe_audio_buffer(self):
        """Транскрибирует весь аудио буфер за раз"""
        try:
            if not self.audio_buffer:
                logger.info("Аудио буфер пуст, нечего транскрибировать")
                self.root.after(0, lambda: self.status_var.set("Нет аудио для транскрибации"))
                self.root.after(0, lambda: self.record_btn.config(text="Начать запись"))
                self.root.after(0, lambda: self.indicator_canvas.itemconfig(self.indicator, fill="gray"))
                return

            # Создаем общий аудио файл из буфера
            combined_wav_data = io.BytesIO()
            
            # Используем первый фрейм для определения параметров
            first_frame = self.audio_buffer[0]
            sample_rate = first_frame['sample_rate']
            sample_width = first_frame['sample_width']
            
            # Проверяем корректность параметров аудио для WAV формата
            if sample_width not in [1, 2, 4]:
                logger.warning(f"Неподдерживаемый sample_width: {sample_width}")
                self.root.after(0, lambda: self.status_var.set("Неподдерживаемый формат аудио"))
                self.root.after(0, lambda: self.record_btn.config(text="Начать запись"))
                self.root.after(0, lambda: self.indicator_canvas.itemconfig(self.indicator, fill="gray"))
                return
            
            if sample_rate <= 0 or sample_rate > 192000:  # Максимальная частота дискретизации для WAV
                logger.warning(f"Неподдерживаемая частота дискретизации: {sample_rate}")
                self.root.after(0, lambda: self.status_var.set("Неподдерживаемый формат аудио"))
                self.root.after(0, lambda: self.record_btn.config(text="Начать запись"))
                self.root.after(0, lambda: self.indicator_canvas.itemconfig(self.indicator, fill="gray"))
                return
            
            # Создаем WAV файл с объединенными данными
            with wave.open(combined_wav_data, 'wb') as wav_file:
                wav_file.setnchannels(1)  # Обычно микрофон - моно (1 канал)
                wav_file.setsampwidth(sample_width)
                wav_file.setframerate(sample_rate)
                
                # Записываем все фреймы
                for frame in self.audio_buffer:
                    # Проверяем, что параметры совпадают с первым фреймом
                    if frame['sample_rate'] != sample_rate or frame['sample_width'] != sample_width:
                        logger.warning("Несоответствие параметров аудио во фрейме, пропускаем")
                        continue
                    wav_file.writeframes(frame['frame_data'])
            
            # Сохраняем временный файл
            with tempfile.NamedTemporaryFile(delete=False, suffix='.wav') as temp_file:
                temp_file.write(combined_wav_data.getvalue())
                temp_filename = temp_file.name
            logger.info(f"Временный файл создан для транскрибации: {temp_filename}")

            # Выполняем транскрибацию
            self.root.after(0, lambda: self.status_var.set("Выполняется транскрибация..."))
            result = self.whisper_model.transcribe(temp_filename, language=self.get_whisper_language_code())
            text = result['text']
            logger.info(f"Транскрибация завершена. Результат: {text}")
            
            # Обновляем интерфейс в основном потоке
            self.root.after(0, self.display_transcription_result, text)
            
        except FileNotFoundError as e:
            # Ошибка может быть вызвана отсутствием ffmpeg или другого необходимого компонента
            logger.error("Ошибка: необходимый компонент не найден. Установите ffmpeg для корректной работы Whisper.", exc_info=True)
            self.root.after(0, lambda: self.status_var.set("Ошибка: отсутствует необходимый компонент (ffmpeg), см. логи..."))
            self.root.after(0, lambda: self.record_btn.config(text="Начать запись"))
            self.root.after(0, lambda: self.indicator_canvas.itemconfig(self.indicator, fill="gray"))
        except Exception as e:
            logger.error(f"Ошибка транскрибации: {e}", exc_info=True)
            self.root.after(0, lambda: self.status_var.set(f"Ошибка транскрибации: {str(e)}"))
            self.root.after(0, lambda: self.record_btn.config(text="Начать запись"))
            self.root.after(0, lambda: self.indicator_canvas.itemconfig(self.indicator, fill="gray"))
        finally:
            # Удаляем временный файл
            try:
                if 'temp_filename' in locals() and os.path.exists(temp_filename):
                    os.unlink(temp_filename)
            except:
                pass  # Игнорируем ошибки при удалении временного файла

    def display_transcription_result(self, text):
        """Отображает результат транскрибации в интерфейсе"""
        # Очищаем текстовое поле и добавляем новый текст
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert(tk.END, text)
        self.text_area.see(tk.END)
        
        # Копируем в буфер обмена
        pyperclip.copy(text)
        
        # Обновляем статус и кнопки
        self.status_var.set("Транскрибация завершена! Результат скопирован в буфер обмена.")
        self.record_btn.config(text="Начать запись")
        self.indicator_canvas.itemconfig(self.indicator, fill="gray")

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
        """Очищает текстовое поле, буфер и буфер обмена"""
        self.text_area.delete("1.0", tk.END)
        self.audio_buffer = []  # Очищаем аудио буфер
        try:
            pyperclip.copy("")  # Очищаем буфер обмена
        except:
            pass  # Если не удалось очистить буфер обмена, игнорируем ошибку
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

    def get_input_devices(self):
        """Получение списка доступных устройств ввода"""
        try:
            audio = pyaudio.PyAudio()
            device_count = audio.get_device_count()
            input_devices = []
            
            for i in range(device_count):
                info = audio.get_device_info_by_index(i)
                if info['maxInputChannels'] > 0:  # Устройство поддерживает ввод
                    # Декодируем имя устройства, чтобы избежать проблем с кодировкой
                    device_name = info['name']
                    try:
                        # Попробуем декодировать строку, если она закодирована в неправильной кодировке
                        if isinstance(device_name, bytes):
                            device_name = device_name.decode('utf-8')
                        else:
                            # Если строка содержит искаженные символы, попробуем преобразовать
                            # Попытка исправить кодировку, если есть проблемы с отображением
                            device_name = device_name.encode('utf-8', errors='ignore').decode('utf-8')
                    except UnicodeDecodeError:
                        # Если возникла ошибка декодирования, используем оригинальное имя
                        device_name = info['name']
                    
                    input_devices.append((i, device_name))
            
            audio.terminate()
            return input_devices
        except Exception as e:
            logger.error(f"Ошибка при получении списка устройств ввода: {e}")
            return []

    def on_device_selected(self, event=None):
        """Обработка выбора устройства ввода"""
        selected_text = self.device_var.get()
        if self.input_devices and selected_text != "Нет доступных устройств":
            # Извлекаем ID устройства из строки (например, "0: Microphone Name" -> 0)
            device_id_str = selected_text.split(':')[0]
            try:
                self.selected_device_id = int(device_id_str)
            except ValueError:
                self.selected_device_id = None
        else:
            self.selected_device_id = None

    def update_sensitivity_label(self, *args):
        """Обновление метки с текущим значением чувствительности"""
        value = int(self.sensitivity_var.get())
        self.sensitivity_label.config(text=str(value))

    def update_energy_threshold(self):
        """Обновление порога чувствительности"""
        self.recognizer.energy_threshold = int(self.sensitivity_var.get())

    def on_model_selected(self, event=None):
        """Обработка выбора модели распознавания"""
        selected_model = self.model_var.get()
        logger.info(f"Выбрана модель распознавания: {selected_model}")
        
        # Загружаем новую модель в отдельном потоке, чтобы не блокировать UI
        threading.Thread(target=self.load_model_async, args=(selected_model,), daemon=True).start()

    def load_model_async(self, model_name):
        """Асинхронная загрузка модели"""
        try:
            logger.info(f"Начинается загрузка модели {model_name}...")
            self.root.after(0, lambda: self.status_var.set(f"Загрузка модели {model_name}..."))
            self.whisper_model = whisper.load_model(model_name)
            logger.info(f"Модель {model_name} успешно загружена")
            self.root.after(0, lambda: self.status_var.set(f"Модель {model_name} загружена"))
        except Exception as e:
            logger.error(f"Ошибка при загрузке модели {model_name}: {e}")
            error_msg = f"Ошибка загрузки модели {model_name}: {str(e)}"
            self.root.after(0, lambda msg=error_msg: self.status_var.set(msg))


def main():
    root = tk.Tk()
    root.title("⏯ Голосовой - ⏯ранскрибер")  # Добавляем символ ⏯ в заголовок окна
    app = VoiceTranscriberApp(root)
    
    # Привязываем клавишу Escape к выходу из программы
    root.bind('<Escape>', lambda e: root.destroy())
    
    root.mainloop()


if __name__ == "__main__":
    main()
