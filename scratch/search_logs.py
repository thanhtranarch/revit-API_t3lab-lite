import json

def search_transcript():
    path = r"C:\Users\trant\.gemini\antigravity-ide\brain\a2af12a4-d6e5-481c-aabc-284d62542b26\.system_generated\logs\transcript.jsonl"
    print("Searching transcript.jsonl...")
    
    with open(path, 'r', encoding='utf-8') as f:
        for line in f:
            try:
                obj = json.loads(line)
                content = obj.get('content', '')
                if 'StackPanel' in content and 'Padding' in content:
                    print("Step {}: Source={}".format(obj.get('step_index'), obj.get('source')))
                    # print snippet
                    print(content[:500])
                    print("-" * 50)
            except Exception as e:
                pass

if __name__ == '__main__':
    search_transcript()
