# 🛡️ Operium Suport
> Central de Controle e Diagnóstico Remoto para Infraestrutura de TI.

![Versão](https://img.shields.io/badge/Versão-5.0_The_Automation_Update-blue)
![Python](https://img.shields.io/badge/Python-3.x-green)
![OS](https://img.shields.io/badge/OS-Windows-lightgrey)

## 📖 Sobre o Projeto

O **Operium Suport** é uma ferramenta centralizada de suporte técnico voltada para analistas e técnicos de infraestrutura de TI corporativa. Desenvolvido para resolver o problema da lentidão no diagnóstico e a fragmentação das ferramentas nativas do Windows, o Operium atua como um painel de controle unificado.

O objetivo do software é permitir que a equipe de suporte localize instantaneamente usuários na rede, realize inventários invisíveis, monitore o desempenho do hardware e execute manutenções remotas avançadas sem interromper o fluxo de trabalho do colaborador.

## ✨ Principais Funcionalidades

* 🌐 **Auto-Load Active Directory:** Varredura automática e assíncrona do AD local, resolvendo nomes DNS e listando computadores online via multithreading (100+ threads simultâneas).
* 🕵️ **Caçador de Usuários (Hunter Module):** Localização em tempo real de usuários na rede a partir do login (via *WMIC probing* de baixa latência).
* 📊 **Monitor de Desempenho Remoto:** Dashboard nativo que espelha o Gerenciador de Tarefas do Windows remotamente, exibindo gráficos de consumo de CPU, RAM e *uptime* em tempo real.
* 📦 **Inventário de Hardware Inteligente:** Extração de dados cruciais via WMI, incluindo IP, MAC, versão do SO, consumo de disco e detecção física avançada (identificação de SSD vs HDD).
* ⚡ **Central de Comandos Remotos:** Disparo simplificado de MSRA (Acesso Remoto), Gpupdate (forçar políticas de grupo), limpeza de Spooler de impressão e renovação de IP local.
* ⚠️ **Stress Test Cirúrgico:** Módulo de reinicialização em loop para testes de bancada, com monitoramento inteligente de ping e *WMI probing* para identificar a prontidão do sistema operacional a cada ciclo.

## 🚀 Pré-requisitos

Para que o Operium funcione com todo o seu potencial, o ambiente deve atender aos seguintes requisitos:
* Sistema Operacional alvo e host executando **Windows 10 ou 11**.
* Máquinas integradas a um domínio **Active Directory** (para as funcionalidades de rede).
* **Privilégios de Administrador** (O script possui elevação automática de UAC ao ser iniciado).
* Serviços WMI, RPC e ICMP (Ping) liberados no Firewall local das máquinas alvo.
* **Python 3.x** instalado.

## 🛠️ Como Executar

1. Clone este repositório para a sua máquina local:
   ```bash
   git clone [https://github.com/SEU_USUARIO/operium-suport.git](https://github.com/SEU_USUARIO/operium-suport.git)
