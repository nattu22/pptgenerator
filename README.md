# SlideDeck AI - Autonomous Presentation Generator

**SlideDeck AI** is an advanced, AI-powered system that autonomously generates professional PowerPoint presentations. It goes beyond simple text-to-slide conversion by orchestrating a multi-agent workflow that plans the narrative, researches content, selects appropriate layouts, and generates visual elements like charts and tables.

## 🚀 Key Features

*   **Intelligent Planning**: Instead of blindly generating slides, the system first creates a comprehensive research plan, determining the optimal number of slides, section topics, and flow based on your query.
*   **Executive Storytelling**: Designed to produce "consulting-style" decks with clear narratives, executive summaries, and structured arguments.
*   **Dynamic Layout Selection**: Uses a smart `TemplateAnalyzer` to inspect your PowerPoint template and select the best layout for each slide's content (e.g., charts for data, columns for comparisons).
*   **Web Search Integration**: Capable of performing simulated or real web searches to gather factual, quantitative data for your slides.
*   **Visual Content Generation**: Automatically generates charts (bar, column, pie) and tables based on the research data.
*   **Interactive Refinement**: A web-based UI allows you to review the plan, edit search queries, and even chat with individual slides to refine content before downloading.
*   **File Upload Support**: Upload PDF, TXT, CSV, or Excel files to generate presentations based on your own documents.

## 🛠️ Architecture

The system is built on a modular architecture:

*   **`flask_app.py`**: The backend API and web server.
*   **`slidedeckai.core`**: The central coordinator (`SlideDeckAI`) for the generation process.
*   **`slidedeckai.agents`**:
    *   `PlanGeneratorOrchestrator`: BREAKS down the user query into a structured plan.
    *   `ExecutionOrchestrator`: EXECUTES the plan, generating content and slides.
    *   `SearchExecutor`: Retrieval agent for gathering facts.
    *   `ContentGenerator`: LLM agent for writing slide text and structuring data.
*   **`slidedeckai.layout_analyzer`**: Intelligent engine that parses PPTX templates to understand available layouts and their suitability.

## 📦 Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/yourusername/slidedeckai.git
    cd slidedeckai
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
    *Note: You may need to install `playwright` separately for testing.*

3.  **Environment Setup:**
    Create a `.env` file in the root directory and add your API keys:
    ```env
    OPENAI_API_KEY=sk-...
    # Optional: Other provider keys if using them
    ```

## 🏃 Usage

### Web Interface (Recommended)

1.  Start the application:
    ```bash
    python flask_app.py
    ```
2.  Open your browser and navigate to `http://localhost:5000`.
3.  **Generate a Deck**:
    *   Enter a topic (e.g., "AI Agents in 2027").
    *   Select a template.
    *   Click "Analyze & Create Plan".
4.  **Refine**:
    *   Review the generated plan. You can edit section titles or search queries.
    *   Click "Approve & Generate Slides".
5.  **Preview & Download**:
    *   Preview the generated slides in the UI.
    *   Use the Chat feature to refine specific slides (e.g., "Make this bullet point shorter").
    *   Download the final `.pptx` file.

### CLI

You can also use the command-line interface for quick generation:

```bash
python src/slidedeckai/cli.py generate --topic "Future of Space Exploration" --model "[oa]gpt-4o" --output-path space_deck.pptx
```

## 🧪 Testing

To verify the system functionality, you can run the provided verification script (requires Playwright):

```bash
pip install playwright
playwright install
python verification_script.py
```

## 📝 Documentation

The codebase is fully documented. Check the source files in `src/slidedeckai` for detailed docstrings on classes and methods.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

MIT License
