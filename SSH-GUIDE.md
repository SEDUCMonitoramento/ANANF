# Guia Rápido de Chaves SSH

Configuração realizada para gerenciar 3 contas GitHub simultaneamente.

## Como Clonar Repositórios

Use o alias correto para cada conta:

| Conta | Alias | Exemplo de Clone |
|-------|-------|------------------|
| **Pessoal** (Gerson-Santiago) | `github.com-pessoal` | `git clone git@github.com-pessoal:Gerson-Santiago/repo.git` |
| **Profissional** (GersonSantiago95) | `github.com-profissional` | `git clone git@github.com-profissional:GersonSantiago95/repo.git` |
| **SEDUC** (SEDUCMonitoramento) | `github.com-seduc` | `git clone git@github.com-seduc:SEDUCMonitoramento/repo.git` |

## Como Configurar Repositórios Já Existentes

Se você já clonou um repositório e quer mudar a conta que ele usa:

1. Abra o terminal na pasta do projeto
2. Execute o comando para alterar a URL remota:

**Para conta Pessoal:**
```bash
git remote set-url origin git@github.com-pessoal:Usuario/Repositorio.git
```

**Para conta Profissional:**
```bash
git remote set-url origin git@github.com-profissional:Usuario/Repositorio.git
```

**Para conta SEDUC:**
```bash
git remote set-url origin git@github.com-seduc:Usuario/Repositorio.git
```

## Testar Conexão

Para verificar se está tudo funcionando:

```bash
ssh -T github.com-pessoal
ssh -T github.com-profissional
ssh -T github.com-seduc
```
