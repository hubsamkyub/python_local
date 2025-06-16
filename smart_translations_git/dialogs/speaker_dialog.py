import tkinter as tk
from tkinter import ttk, messagebox
from scenario_manager import SpeakerProfile

class SpeakerEditDialog:
    """화자 편집 다이얼로그"""
    
    def __init__(self, parent, speaker=None, callback=None):
        self.top = tk.Toplevel(parent)
        self.top.title("화자 편집" if speaker else "새 화자 추가")
        self.top.geometry("600x500")
        self.top.transient(parent)
        self.top.grab_set()
        
        self.speaker = speaker
        self.callback = callback
        
        self.setup_ui()
        
        if speaker:
            self.load_speaker_data()
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.top, padding="15")
        main_frame.pack(fill="both", expand=True)
        
        # 기본 정보
        basic_frame = ttk.LabelFrame(main_frame, text="기본 정보")
        basic_frame.pack(fill="x", pady=5)
        
        # 이름
        ttk.Label(basic_frame, text="화자 이름:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.name_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=self.name_var, width=30).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        # 성별
        ttk.Label(basic_frame, text="성별:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.gender_var = tk.StringVar(value="중성")
        gender_combo = ttk.Combobox(basic_frame, textvariable=self.gender_var, values=["남성", "여성", "중성"], state="readonly")
        gender_combo.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        # 말투
        ttk.Label(basic_frame, text="말투:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.tone_var = tk.StringVar(value="보통")
        tone_combo = ttk.Combobox(basic_frame, textvariable=self.tone_var, values=["정중", "보통", "친근", "거친", "우아", "유치"], state="readonly")
        tone_combo.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        basic_frame.grid_columnconfigure(1, weight=1)
        
        # 스타일 설명
        style_frame = ttk.LabelFrame(main_frame, text="번역 스타일 설명")
        style_frame.pack(fill="x", pady=5)
        
        ttk.Label(style_frame, text="이 화자의 번역 특징을 설명하세요:").pack(anchor="w", padx=5, pady=2)
        self.style_text = tk.Text(style_frame, height=3, wrap="word")
        self.style_text.pack(fill="x", padx=5, pady=5)
        
        # 예시 문장
        example_frame = ttk.LabelFrame(main_frame, text="번역 예시")
        example_frame.pack(fill="both", expand=True, pady=5)
        
        ttk.Label(example_frame, text="이 화자의 번역 예시를 추가하세요 (KR → EN):").pack(anchor="w", padx=5, pady=2)
        
        # 예시 입력 프레임
        add_example_frame = ttk.Frame(example_frame)
        add_example_frame.pack(fill="x", padx=5, pady=2)
        
        ttk.Label(add_example_frame, text="KR:").grid(row=0, column=0, sticky="w")
        self.kr_example_var = tk.StringVar()
        ttk.Entry(add_example_frame, textvariable=self.kr_example_var, width=40).grid(row=0, column=1, sticky="ew", padx=2)
        
        ttk.Label(add_example_frame, text="EN:").grid(row=1, column=0, sticky="w")
        self.en_example_var = tk.StringVar()
        ttk.Entry(add_example_frame, textvariable=self.en_example_var, width=40).grid(row=1, column=1, sticky="ew", padx=2)
        
        ttk.Button(add_example_frame, text="예시 추가", command=self.add_example).grid(row=0, column=2, rowspan=2, padx=5)
        
        add_example_frame.grid_columnconfigure(1, weight=1)
        
        # 예시 목록
        self.example_listbox = tk.Listbox(example_frame, height=6)
        self.example_listbox.pack(fill="both", expand=True, padx=5, pady=5)
        
        example_btn_frame = ttk.Frame(example_frame)
        example_btn_frame.pack(fill="x", padx=5)
        ttk.Button(example_btn_frame, text="선택 삭제", command=self.delete_example).pack(side="left")
        
        # 버튼
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        ttk.Button(button_frame, text="취소", command=self.cancel).pack(side="right", padx=5)
        ttk.Button(button_frame, text="저장", command=self.save).pack(side="right")
    
    def load_speaker_data(self):
        """기존 화자 데이터 로드"""
        if not self.speaker:
            return
            
        self.name_var.set(self.speaker.name)
        self.gender_var.set(self.speaker.gender)
        self.tone_var.set(self.speaker.tone)
        self.style_text.insert("1.0", self.speaker.style)
        
        for example in self.speaker.examples:
            self.example_listbox.insert("end", f"KR: {example['kr']} | EN: {example['en']}")
    
    def add_example(self):
        """예시 추가"""
        kr = self.kr_example_var.get().strip()
        en = self.en_example_var.get().strip()
        
        if kr and en:
            self.example_listbox.insert("end", f"KR: {kr} | EN: {en}")
            self.kr_example_var.set("")
            self.en_example_var.set("")
        else:
            messagebox.showwarning("입력 오류", "KR과 EN을 모두 입력하세요.")
    
    def delete_example(self):
        """선택된 예시 삭제"""
        selection = self.example_listbox.curselection()
        if selection:
            self.example_listbox.delete(selection[0])
    
    def save(self):
        """화자 정보 저장"""
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("입력 오류", "화자 이름을 입력하세요.")
            return
        
        # 예시 데이터 수집
        examples = []
        for i in range(self.example_listbox.size()):
            line = self.example_listbox.get(i)
            parts = line.split(" | ")
            if len(parts) == 2:
                kr_part = parts[0].replace("KR: ", "")
                en_part = parts[1].replace("EN: ", "")
                examples.append({"kr": kr_part, "en": en_part})
        
        # 화자 객체 생성
        speaker = SpeakerProfile(
            name=name,
            gender=self.gender_var.get(),
            tone=self.tone_var.get(),
            style=self.style_text.get("1.0", "end-1c").strip(),
            examples=examples
        )
        
        if self.callback:
            self.callback(speaker)
        
        self.top.destroy()
    
    def cancel(self):
        """취소"""
        self.top.destroy()