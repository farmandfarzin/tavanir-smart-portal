# Tavanir Smart Portal

An enterprise-grade, portable data analytics and intelligent decision-support portal designed for managing and auditing large-scale energy consumption data (electricity, water, and gas). Built on top of Streamlit and powered by advanced AI and data visualization tools, this platform enables robust institutional auditing and energy optimization.

### Key Features
* **AI-Powered Analytics:** Integrated with local/offline LLM infrastructure and RAG (Retrieval-Augmented Generation) for secure, intelligent data querying.
* **Dynamic Dashboards:** Interactive data visualizations powered by Plotly for tracking and forecasting utility usage across facilities.
* **Role-Based Authentication:** Secure user management and credential verification using `streamlit-authenticator` and YAML configuration.
* **100% Portable:** Fully decoupled from fixed directory structures; packaged using PyInstaller to run seamlessly on any target machine without external dependencies or pre-installed Python environments.

### Quick Start
To run the portable production version:
1. Extract the `run_app` bundle on the target system.
2. Double-click `run_app.exe` to launch the native server.
3. The secure portal will automatically open in your default browser at `http://localhost:8501`.

### Tech Stack
* **Frontend/Server:** Streamlit
* **AI Framework:** LangChain, Torch, Transformers
* **Data & Plots:** Pandas, OpenPyXL, Plotly
* **Security:** Streamlit Authenticator (YAML-based)
