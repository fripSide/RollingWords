# -*- coding: utf-8 -*-
import io
import os
import sys
import json
import codecs
import time
import threading
import xlrd
from colorama import init, Style, Fore
from getkey import getkey, keys

__version__ = "2019-7-18 version: 0.9"

IS_PY3 = sys.version_info > (3, 3)
IS_WIN = os.name == "nt"
# sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8') 

if not IS_PY3:
	reload(sys)
	sys.setdefaultencoding("utf-8")

osp = os.path
init(autoreset=True)

def get_str_bytes(s):
	# 从excel读取的内容（由于系统编码被设置为utf-8，所以内容为utf-8）
	try:
		return s.decode("utf-8")
	except:
		return s

def convert_to_utf8(s):
	# 从系统读取的文件名,例如:路径名
	try:
		if IS_WIN:
			return s.decode("gbk", "ignore")
		else:
			return s.decode("utf-8")
	except:
		return s.decode("utf-8")

def get_print_str(s):
	# 将编码格式转成跟当前的stdout一致，避免print报错
	by = get_str_bytes(s)
	pw = by.encode(sys.stdout.encoding, "ignore")
	return pw

def print_log(msg, *args):
	if IS_PY3:
		print(msg, args)
	else:
		s = u"".join([str(a) for a in args])
		print msg, s

class ExcelBookLoader:
	"""
	暂时只支持Excel格式：第一列单词，第二列意义。
	"""
	def __init__(self, book_path):
		self.book_path = book_path
		self.sheet = None

	def load_book(self):
		all_words = []
		try:
			excel = xlrd.open_workbook(self.book_path)
			self.sheet = excel.sheet_by_index(0)
			for i in range(0, self.sheet.nrows):
				all_words.append(self._get_item(i))
		except Exception as ex:
			print_log(Fore.RED + u"Failed to load: {} {}".format(self.book_path, ex.args))
			return []
		return all_words

	def _get_item(self, idx):
		word = self.sheet.cell(rowx=idx, colx=0).value
		meaning = self.sheet.cell(rowx=idx, colx=1).value
		word = word.strip()
		word = get_str_bytes(word)
		meaning = "- " + meaning.replace(u'；', '\n- ').replace(u'; ', '\n- ').replace(u';', '\n- ')
		meaning = get_str_bytes(meaning)
		return [word, meaning]

class BookDatabase:
	RECORD = "record.json"
	MODE_NORMAL = "all words"
	MODE_HARD = "hard words"
	MODES = [MODE_NORMAL, MODE_HARD]

	def __init__(self, reload_book=False, book_dir="books"):
		self.reload_book = reload_book
		self.book_dir = book_dir
		self.need_eixt = False
		self.cur_book = None
		self.all_words = None
		self.word_list = []
		self.normal_pos = 0
		self.cur_pos = 0
		self.hard_pos = 0
		self.mode = self.MODE_NORMAL
		self.hard_list = []
		self.record = {}
		self.read_lock = threading.Lock()

	def load_book(self):
		load_history = False
		self.need_eixt = True
		load_history = self._load_history()
		if load_history and self.reload_book:
			print_log(u"Current book is {}. Do you want to change book(y/n)?".format(self.cur_book))
			c = check_and_getkey(['y', 'n', 'Y', 'N', keys.ENTER])
			if c == 'y' or c == 'Y':
				load_history = False
			elif c == 'e':
				return
		if not load_history:
			self._choose_book()
		if self.cur_book:
			self.all_words = self._load_all_words(self.cur_book)
		if not self.all_words:
			return
		self._choose_mode()

	def _load_history(self):
		# 读取当前书的学习进度
		if osp.exists(self.RECORD):
			self.record = self._load_json(self.RECORD)
			if not self.cur_book:
				book_name = self.record.get("current")
			else:
				book_name = self.cur_book
			books = self.record.get("books", {})
			if book_name in books:
				book = books.get(book_name, {})
				self.cur_book = book_name
				self.normal_pos = book.get("normal_pos", 0)
				self.hard_pos = book.get("hard_pos", 0)
				self.hard_list = book.get("hard_list", [])
				return True
		return False

	def _load_all_words(self, book_name):
		# 暂时只支持excel
		loader = ExcelBookLoader(osp.join(self.book_dir, book_name))
		all_words = loader.load_book()
		if all_words:
			print_log(u'Current book: {}'.format(book_name))
		return all_words

	def _load_json(self, path):
		try:
			with codecs.open(path, "r", "utf-8") as fp:
				return json.load(fp, encoding="utf-8")
		except Exception as ex:
			print_log(Fore.YELLOW + u"Faild to load record: {}".format(ex.args))
			return {}

	def _choose_book(self):
		self.cur_book = None
		books = os.listdir(self.book_dir)
		idx = 1
		print_log("Please choose a book:")
		options = {}
		for bk in books:
			if bk.endswith(".xls") or bk.endswith(".xlsx"):
				print_log("> [{}] {}".format(idx, bk))
				options[str(idx)] = convert_to_utf8(bk)
				idx += 1
		c = check_and_getkey(list(options.keys()))
		if c in options:
			self.cur_book = options[c]
			self._load_history()

	def _choose_mode(self):
		print_log(u"Roll all words or roll hard words(1/2)?")
		c = check_and_getkey(['1', '2'])
		if c == 'e': return
		self.mode = self.MODES[int(c) - 1]
		self.need_eixt = False
		if self.mode == self.MODE_NORMAL:
			self.cur_pos = self.normal_pos
			self.word_list = [i for i in range(0, len(self.all_words))]
		elif self.mode == self.MODE_HARD:
			self.cur_pos = self.hard_pos
			self.word_list = self.hard_list
		if self.cur_pos >= len(self.word_list):
			self.cur_pos = 0

	def get_word(self, idx):
		if idx >= 0 and idx < len(self.word_list):
			word_idx = self.word_list[idx]
			return self.all_words[word_idx]
		return None

	def next_word(self):
		with self.read_lock:
			word = self.get_word(self.cur_pos)
			if not word:
				print("All words in list are finished [{}].".format(len(self.word_list)))
				return
			self.cur_pos += 1
			return word

	def mark_hard(self, add):
		with self.read_lock:
			cur = self.cur_pos - 1
			item = self.get_word(cur)
			word = item[0]
			if add:
				print(u"=> Mark {} as hard".format(word))
				if cur not in self.hard_list:
					self.hard_list.append(cur)
			else:
				print(u"=> Delete {} from hard list".format(word))
				del self.hard_list[cur]

	def get_pos_range(self):
		return self.cur_pos - 1, len(self.word_list)

	@property
	def pos(self):
		total_num = len(self.word_list)
		self.cur_pos = self.cur_pos % total_num
		return self.cur_pos

	def save_results(self):
		if self.cur_book and self.cur_pos > 0:
			self.record["current"] = self.cur_book
			self.record["books"] = self.record.get("books", {})
			books = self.record["books"]
			books[self.cur_book] = books.get(self.cur_book, {})
			cur_book = books[self.cur_book]
			if self.mode == self.MODE_NORMAL:
				cur_book["normal_pos"] = self.cur_pos
			if self.mode == self.MODE_HARD:
				cur_book["hard_pos"] = self.cur_pos
			cur_book["hard_list"] = self.hard_list
			with codecs.open(self.RECORD, "w", "utf-8") as fp:
				json.dump(self.record, fp, ensure_ascii=False, indent=4)

class CmdControl:
	TI_WAIT = 1
	TI_SHOW = 2

	def __init__(self, change_book):
		self.db = BookDatabase(change_book)
		self.db.load_book()
		self.audio_engine = None
		self.play_audio = True
		self.pronouce_list = []
		self.should_exit = self.db.need_eixt
		self.word_thread = None
		self.speech_thread = None
		self.is_pause = False
		self.show_cur = True # 是否持续显示，跳过的时候设为False
		self._start_to_learn()

	def _start_to_learn(self):
		if self.should_exit: return
		self.speech_thread = threading.Thread(target=self.word_pronouse)
		self.speech_thread.setDaemon(True)
		self.speech_thread.start()
		self.word_thread = threading.Thread(target=self.word_display)
		self.word_thread.setDaemon(True)
		self.word_thread.start()
	

	def word_pronouse(self):
		# Pyttsx需要在同一个线程中初始化和调用runAndWait，不然runAndWait会无法返回
		try:
			import pyttsx
			self.audio_engine = pyttsx.init()
		except Exception as ex:
			print_log(Fore.RED + u"Faild to init speech engine: {}".format(ex.args))
			self.play_audio = False
			return
		while not self.should_exit:
			# print_log("wait to speak before", len(self.pronouce_list))
			if len(self.pronouce_list) > 0:
				text = self.pronouce_list.pop()
				if not self.text_to_speech(text):
					break
			time.sleep(0.5)
		# print_log("Speech_thread threading exit!")
		self.play_audio = False

	def text_to_speech(self, text):
		""" https://pythonprogramminglanguage.com/text-to-speech/
		import win32com.client
		speaker = win32com.client.Dispatch("SAPI.SpVoice")
		speaker.Speak(text)
		"""
		try:
			# print("want to speak ", text)
			text = text.split(" ")[0]
			self.audio_engine.say(text)
			self.audio_engine.runAndWait()
			# print("finish ", text)
		except Exception as ex:
			print_log(Fore.RED + u"Faild to init speech engine: {}".format(ex.args))
			return False
		return True

	def word_display(self):
		item = self.db.next_word()
		while item and not self.should_exit:
			self.check_pause()
			self.print_one_word(item)
			item = self.db.next_word()
		# print_log("Word_thread threading exit!")
		self.should_exit = True
		print_log("Task is finished and press any key to exit!")

	def _save_results(self):
		print_log(Fore.GREEN + "Save and exit!")
		self.db.save_results()

	def print_one_word(self, item):
		self.show_cur = True
		c, t, = self.db.get_pos_range()
		os.system('cls' if IS_WIN else 'clear')
		print("")
		fraction = int((float(c) / t) * 40)
		fraction_str = "[%d/%d]" % (c, t)
		print("=" * fraction + "." * (40 - fraction) + fraction_str)
		word, meaning = item[0], item[1]
		text = get_print_str(word)
		print(u"=>" + "[{}]. ".format(self.db.cur_pos) + Fore.GREEN + text + Fore.RESET)
		if self.play_audio:
			self.pronouce_list.append(word)
		self.show_in_time(self.TI_WAIT)
		if not self.show_cur:
			return
		print(get_print_str(meaning))
		self.show_in_time(self.TI_SHOW)

	def show_in_time(self, ti):
		while ti > 0 and self.show_cur:
			time.sleep(0.1)
			ti -= 0.1

	def check_pause(self):
		while self.is_pause:
			os.system('cls' if IS_WIN else 'clear')
			print("pause!")
			time.sleep(0.1)

	def wait_to_exit(self):
		self._save_results()
		self.should_exit = True
		exit(0)
		# self.word_thread.join()
		# if self.speech_thread:
			# self.speech_thread.join()

def check_and_getkey(options=None):
	while True:
		try:
			c = getkey()
		except KeyboardInterrupt as ex:
			raise KeyboardInterrupt
			c = 'e'
		if not options or c == 'e': 
			return c
		if c not in options:
			print_log(Fore.RED + "Need to input {}!".format(options))
			continue
		else:
			return c

def run_loop(change_book=False):
	ctrl = CmdControl(change_book)
	while True:
		if ctrl.should_exit:
			break
		c = check_and_getkey()
		# print_log(u"input: {}".format(c))
		if c == 'e':
			break
		elif c == 'p':
			ctrl.is_pause = not ctrl.is_pause
		elif c == 'h':
			ctrl.mark_hard(True)
		elif c == 'd':
			ctrl.mark_hard(False)
		elif c == 'n':
			ctrl.show_cur = False
		elif c == 'b':
			ctrl.db.cur_pos -= 2
			ctrl.show_cur = False
	ctrl.wait_to_exit()

def print_usage():
	print_log("| Automatic vocabulary roller, by fripSide")
	print_log("| Modified from Atlantix_Vocabulary_Roller.py")
	print_log("| {}".format(__version__))
	print_log("| Usage:")
	print_log("=> " + Fore.GREEN + "[e]xit" + Fore.RESET + " [p]ause")
	print_log("=> mark [h]ard, or [d]elete")
	print_log("=> [n]ext, or [b] for previous")

def main():
	change_book = False
	if len(sys.argv) > 1:
		print(sys.argv)
		if sys.argv[1] == '-c':
			change_book = True
	print_usage()
	run_loop(change_book)

if __name__ == "__main__":
	main()