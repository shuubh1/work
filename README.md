# Work

<!-- Add a brief description here. For example: -->
> An interactive web application built with Streamlit.

## 🚀 Features

- **Interactive Dashboard**: Main interface driven by `Home.py`.
- **Multi-Page Architecture**: Navigate through different modules via the sidebar (located in the `pages/` directory).
- **Configurable**: Project settings and parameters are managed via `config.yaml`.
- **Modular Design**: Core logic is separated into `utils/` and presentation elements in `templates/`.

## 🛠️ Installation

1. **Clone the repository**
   ```bash
   git clone [https://github.com/shuubh1/work.git](https://github.com/shuubh1/work.git)
   cd work
   ```

2. **Set up a virtual environment (Recommended)**
   ```bash
   # MacOS/Linux
   python3 -m venv venv
   source venv/bin/activate

   # Windows
   python -m venv venv
   .\venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

## ▶️ Usage

To launch the application locally, run the following command from the root directory:
```bash
streamlit run Home.py
```
The app should open automatically in your default browser at `http://localhost:8501`.

## 📂 Project Structure

```plaintext
work/
├── Home.py             # Main entry point for the Streamlit app
├── config.yaml         # Configuration file for app settings
├── pages/              # Directory for additional app pages
├── templates/          # HTML templates or Prompt templates
├── utils/              # Helper functions and utility scripts
├── requirements.txt    # Python package dependencies
└── packages.txt        # System-level dependencies (for deployment)
```

## ☁️ Deployment

This project includes a `packages.txt` file, which suggests it is ready for deployment on Streamlit Community Cloud.
1. Push your changes to GitHub.
2. Log in to [Streamlit Community Cloud](https://streamlit.io/cloud).
3. Connect your GitHub account and select this repository.
4. Set the "Main file path" to `Home.py`.

## 🤝 Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
