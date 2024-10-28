import speech_recognition as sr
import subprocess

# Função para abrir aplicativos
def abrir_programa(programa):
    try:
        if programa == "excel":
            subprocess.Popen("excel.exe")  # Abre o Microsoft Excel
            print("Abrindo Excel...")
        elif programa == "word":
            subprocess.Popen("winword.exe")  # Abre o Microsoft Word
            print("Abrindo Word...")
        elif programa == "outlook":
            subprocess.Popen("outlook.exe")  # Abre o Microsoft Outlook
            print("Abrindo Outlook...")
        else:
            print("Comando não reconhecido.")
    except FileNotFoundError:
        print(f"{programa.capitalize()} não encontrado. Verifique se você possui o programa.")

# Função principal para escutar o comando de voz
def escutar_comando():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Diga o comando para abrir um aplicativo (Ex.: 'abrir excel')")
        audio = recognizer.listen(source)

        try:
            comando = recognizer.recognize_google(audio, language="pt-BR").lower()
            print(f"Você disse: {comando}")
            
            # Verifica o comando e chama a função correspondente
            if "abrir excel" in comando:
                abrir_programa("excel")
            elif "abrir word" in comando:
                abrir_programa("word")
            elif "abrir outlook" in comando:
                abrir_programa("outlook")
            else:
                print("Comando não reconhecido. Tente novamente.")
        except sr.UnknownValueError:
            print("Não entendi o que você disse.")
        except sr.RequestError as e:
            print(f"Erro ao conectar ao serviço de reconhecimento de voz: {e}")

# Loop principal para continuar ouvindo até que o programa seja encerrado manualmente
while True:
    escutar_comando()
