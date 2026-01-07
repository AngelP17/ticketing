Automated Infrastructure & Data Pipeline (IT Ops)

A production-grade, containerized systems monitoring dashboard designed to bridge the gap between legacy manufacturing data (Excel) and modern relational database architectures.

This project demonstrates a full CI/CD pipeline, Docker orchestration, and an automated ETL (Extract, Transform, Load) process.

🏗️ System Architecture

The application is built using a microservices-inspired architecture to ensure modularity and scalability:

Frontend/API: Flask (Python) serving a responsive monitoring dashboard.

Database: SQLite (Relational) for high-speed, indexed data retrieval.

ETL Engine: A background Python service using Pandas/SQLAlchemy that monitors local file changes and syncs data from Excel to SQL.

Reverse Proxy: Nginx (Production-ready traffic routing).

Orchestration: Docker Compose for environment parity across Windows, macOS, and Linux.

🛠️ Engineering Features

1. Automated ETL Pipeline

Instead of manual data entry, I engineered a "Watcher" service that:

Extracts: Monitors tickets.xlsx for updates.

Transforms: Validates data integrity, cleans NaNs, and formats timestamps using Pandas.

Loads: Performs atomic updates to the SQLite relational database.

Outcome: Reduced data latency and eliminated manual sync errors.

2. DevOps & Containerization

The entire stack is containerized using Docker. This ensures:

Environment Parity: The app runs identically in local development and cloud production.

Infrastructure as Code (IaC): Server configuration is codified in docker-compose.yml and Dockerfile.

3. CI/CD Workflow

GitOps: Pushes to the main branch trigger automated builds and deployments.

Cloud Native: Deployed on Render with environment variable management for secure credential handling.

🚦 Getting Started

The "Cloud-Native" Way (Recommended)

Ensure you have Docker Desktop installed.

git clone [https://github.com/your-username/your-repo-name.git](https://github.com/your-username/your-repo-name.git)
cd your-repo-name
docker-compose up --build


Access the app at http://localhost

The Legacy Way (Local Python)

If you prefer running without Docker:

Setup: python -m venv venv

Activate: source venv/bin/activate (Mac) or venv\Scripts\activate (Windows)

Install: pip install -r requirements.txt

Run: python app.py

🧰 Tech Stack

Language: Python 3.10

Framework: Flask

Data Science: Pandas, NumPy

Database: SQLite, SQLAlchemy

DevOps: Docker, Docker Compose, Nginx, GitHub Actions/Hooks

Cloud: Render

👨‍💻 Author

Angel Pinzon Computer Engineer (B.S.Cp.E.) Portfolio | LinkedIn
