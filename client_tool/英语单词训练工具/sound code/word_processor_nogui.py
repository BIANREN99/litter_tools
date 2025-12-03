import json
import csv
from pathlib import Path

class WordProcessor:
    def __init__(self):
        self.words = []
    
    def parse_line(self, line):
        """解析单行单词数据"""
        line = line.strip()
        if not line:
            return None
            
        # 分离单词和释义
        parts = line.split(' ', 1)
        if len(parts) < 2:
            return None
            
        word = parts[0]
        definition = parts[1]
        
        # 分离词性和释义
        if '.' in definition:
            pos, meaning = definition.split('.', 1)
            return {
                'word': word,
                'pos': pos.strip(),
                'meaning': meaning.strip()
            }
        return None
    
    def load_from_file(self, file_path):
        """从文件加载单词数据"""
        self.words = []
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    parsed = self.parse_line(line)
                    if parsed:
                        self.words.append(parsed)
            return True
        except Exception as e:
            print(f"Error loading file: {e}")
            return False
    
    def to_markdown(self, output_path):
        """转换为Markdown表格格式"""
        markdown = "| 单词 | 词性 | 释义 |\n"
        markdown += "|------|------|------|\n"
        
        for word in self.words:
            markdown += f"| {word['word']} | {word['pos']} | {word['meaning']} |\n"
        
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(markdown)
            return True
        except Exception as e:
            print(f"Error saving markdown: {e}")
            return False
    
    def to_csv(self, output_path):
        """转换为CSV格式"""
        try:
            with open(output_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=['word', 'pos', 'meaning'])
                writer.writeheader()
                writer.writerows(self.words)
            return True
        except Exception as e:
            print(f"Error saving CSV: {e}")
            return False
    
    def to_json(self, output_path):
        """转换为JSON格式"""
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(self.words, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"Error saving JSON: {e}")
            return False
    
    def to_txt(self, output_path, separator='|'):
        """转换为带分隔符的文本格式"""
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                for word in self.words:
                    f.write(f"{word['word']}{separator}{word['pos']}{separator}{word['meaning']}\n")
            return True
        except Exception as e:
            print(f"Error saving TXT: {e}")
            return False

def main():
    # 创建处理器实例
    processor = WordProcessor()
    
    # 从文件读取数据
    input_file = input("请输入文件路径（支持.txt和.csv格式): ")
    if processor.load_from_file(input_file):
        print(f"成功从 {input_file} 加载 {len(processor.words)} 个单词")
        
        # 创建输出目录
        output_dir = Path('output')
        output_dir.mkdir(exist_ok=True)
        
        # 转换为不同格式
        processor.to_markdown(output_dir / 'words.md')
        processor.to_csv(output_dir / 'words.csv')
        processor.to_json(output_dir / 'words.json')
        processor.to_txt(output_dir / 'words.txt')
        
        print("转换完成!文件已保存在output目录中。")
    else:
        print("加载文件失败")


if __name__ == '__main__':
    main()
