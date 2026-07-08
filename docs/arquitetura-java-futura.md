# Estrutura do Projeto Java (Spring Boot + Apache POI)

Esta é a estrutura padrão Maven que utilizaremos para o novo Orquestrador Centralizado.

## 1. Estrutura de Pastas (Maven Standard)

```text
orquestrador-swaps/
├── pom.xml                        # Gerenciador de dependências (Apache POI, Spring Boot)
├── src/
│   ├── main/
│   │   ├── java/
│   │   │   └── com/empresa/orquestrador/
│   │   │       ├── OrquestradorApplication.java   # Ponto de entrada (Main)
│   │   │       ├── controller/                   # Interface Web (APIs)
│   │   │       │   └── DashboardController.java
│   │   │       ├── service/                      # Lógica de Negócio (O cérebro)
│   │   │       │   ├── ExcelService.java         # Manipulação via Apache POI
│   │   │       │   └── SftpService.java          # Download do INETX
│   │   │       └── model/                        # Definição dos metadados
│   │   │           └── Calculadora.java
│   │   └── resources/
│   │       ├── application.properties            # Configurações (IP, Portas, Paths)
│   │       ├── static/                           # Front-end (CSS, JS, Imagens)
│   │       └── templates/                        # Front-end (HTML - Thymeleaf)
│   │           └── dashboard.html
└── target/                                # Onde o .JAR executável será gerado
```

## 2. Dependências Principais (pom.xml)

As bibliotecas essenciais que o Java usará para substituir o VBA:

*   **spring-boot-starter-web**: Para criar o servidor e o dashboard no navegador.
*   **poi-ooxml**: A biblioteca para ler e escrever nos arquivos `.xlsx`.
*   **jsch**: Para conectar no SFTP do INETX de forma nativa.
*   **lombok**: Para reduzir o código repetitivo (Getters/Setters).

## 3. Modelo de Dados (Calculadora.java)

Este é o "Metadado" que permite adicionar novas calculadoras sem mexer no código:

```java
public class Calculadora {
    private String nome;
    private String caminhoArquivo;
    private String abaMtm;
    private String celulaData;   // ex: "N6"
    private String celulaCheck;  // ex: "O2"
    private List<String> curvas; // ["Pre", "CDI", "SOFR"]
}
```

## 4. Próximos Passos de Implementação

1.  **Criação do Projeto**: Usar o Spring Initializr para gerar a base.
2.  **Módulo de Leitura**: Criar o `ExcelService` para abrir um arquivo e ler o valor da célula de Check.
3.  **Módulo de Escrita**: Implementar a função que injeta a data e a taxa CDI.
4.  **Interface**: Montar o HTML do Dashboard para exibir a lista de calculadoras.
