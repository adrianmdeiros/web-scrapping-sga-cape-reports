## Propriedades que podem ser úteis para quando se tem um WebDriverElement

value: Retorna o valor do atributo value do elemento, útil em campos de entrada (<input>).

tag_name: Retorna o nome da tag HTML do elemento, como div, input, button, etc.

id: Retorna o valor do atributo id do elemento.

visible: Retorna True se o elemento estiver visível, False caso contrário.

outer_html: Retorna o HTML completo do elemento, incluindo a tag de abertura e fechamento.

html ou inner_html: Retorna o HTML interno do elemento, ou seja, o conteúdo entre a tag de abertura e a de fechamento.

click(): Clica no elemento.
fill(value): Preenche um campo de entrada (<input>, <textarea>) com o valor especificado.
clear(): Limpa o conteúdo do elemento de entrada.
check() e uncheck(): Marcam ou desmarcam uma checkbox.
select(text): Seleciona uma opção em um elemento <select>.
is_selected(): Verifica se um elemento, como um checkbox ou uma opção de seleção, está selecionado.
has_class(class_name): Verifica se o elemento possui uma determinada classe CSS.