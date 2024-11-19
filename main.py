import questionary
import subprocess

file_aliases = {
    'executed_services.py': 'Serviços executados',
    'finished_services.py': 'Atendimentos conluídos',
}

def main():
    aliases = list(file_aliases.values())

    selected_alias = questionary.select(
        "Qual script você quer rodar?",
        choices=aliases
    ).ask() 

    if selected_alias:
        for file, alias in file_aliases.items():
            if alias == selected_alias:
                try:
                    subprocess.run(['python', file], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Erro ao executar {file}: {e}")
                except FileNotFoundError:
                    print(f"O arquivo {file} não foi encontrado!")
                break

if __name__ == "__main__":
    main()