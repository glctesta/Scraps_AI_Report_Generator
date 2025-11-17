"""
AI Analyzer for Quality Analysis
Uses Ollama to generate root cause analysis, recommendations, and Kaizen proposals.
"""
import json
import logging
import requests
from typing import Dict, List, Any

# Setup logger
logger = logging.getLogger('AIAnalyzer')


class AIAnalyzer:
    """Analyzes defect data using an AI model via Ollama."""

    def __init__(self, base_url: str = 'http://localhost:11434', model: str = 'llama3.2:latest'):
        """
        Initializes the AI Analyzer.

        Args:
            base_url (str): The base URL of the Ollama server API.
            model (str): The name of the model to use for analysis.
        """
        if not base_url:
            raise ValueError("Ollama base_url cannot be empty.")
        self.base_url = base_url
        self.model = model
        logger.info(f"AIAnalyzer initialized for model '{self.model}' at {self.base_url}")

    def _call_ai(self, prompt: str) -> Dict | None:
        """Generic method to call the Ollama API and parse the JSON response."""
        try:
            logger.info("Sending request to AI model...")
            response = requests.post(
                f"{self.base_url}/api/generate",
                json={"model": self.model, "prompt": prompt, "stream": False, "format": "json"},
                timeout=300  # 5 minutes for complex analysis
            )
            response.raise_for_status()
            ai_response_str = response.json().get('response', '{}')
            parsed_json = json.loads(ai_response_str)
            logger.info("Successfully received and parsed AI response.")
            return parsed_json
        except requests.exceptions.RequestException as e:
            logger.error(f"Ollama API request failed: {e}")
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse JSON from AI response: {e}. Response was: {ai_response_str[:200]}...")
        except Exception as e:
            logger.error(f"An unexpected error occurred during AI call: {e}", exc_info=True)
        return None

    def analyze_defects(self, top_defects: List[Dict], production_data: Dict) -> Dict:
        """Analyzes scrap defects."""
        prompt = self._create_scrap_analysis_prompt(top_defects, production_data)
        ai_response = self._call_ai(prompt)
        if ai_response:
            return ai_response
        return {'executive_summary': 'AI analysis for scraps failed. Please review statistical data.', 'root_causes': [], 'recommendations': []}

    def analyze_fails(self, fail_data: List[Dict], statistics: Dict, period_type: str) -> Dict:
        """Analyzes production fails and requests a Kaizen proposal."""
        prompt = self._create_fail_analysis_prompt(statistics, period_type)
        ai_response = self._call_ai(prompt)
        if ai_response:
            return ai_response
        return {'executive_summary': 'AI analysis for fails failed. Please review statistical data.', 'root_causes': [], 'recommendations': [], 'kaizen_project_proposal': None}

    def analyze_breakdowns(self, statistics: Dict, production_data: Dict, period_type: str) -> Dict:
        """Analyzes line stoppages and requests a Kaizen proposal."""
        prompt = self._create_stoppage_analysis_prompt(statistics, production_data, period_type)
        ai_response = self._call_ai(prompt)
        if ai_response:
            return ai_response
        return {'executive_summary': 'AI analysis for breakdowns failed. Please review statistical data.', 'root_causes': [], 'recommendations': [], 'kaizen_project_proposal': None}


    # --- PROMPT CREATION METHODS ---

    def _create_scrap_analysis_prompt(self, top_defects: List[Dict], production_data: Dict) -> str:
        defects_summary = "\n".join([f"- {d['DefectName']}: {d['Count']} times" for d in top_defects[:5]])
        total_defects = sum(d['Count'] for d in top_defects)
        nr_boards = production_data.get('NrBoards', 1)
        scrap_rate = (total_defects / nr_boards * 100) if nr_boards > 0 else 0

        return f"""
        You are a Quality Assurance expert in an electronics manufacturing plant.
        Analyze the following weekly scrap data. Your response MUST be a single, valid JSON object and in English.

        **Production Context:**
        - Total Boards Produced: {production_data.get('NrBoards', 'N/A')}
        - Total Scrapped Boards: {total_defects}
        - Weekly Scrap Rate: {scrap_rate:.2f}%

        **Top 5 Defects:**
        {defects_summary}
        
        **Task:**
        Provide a concise analysis in the following JSON format.

        {{
          "executive_summary": "A brief paragraph summarizing the key findings and the overall quality performance for the week.",
          "root_causes": [
            {{
              "problem_area": "e.g., Soldering Process",
              "cause_description": "A probable root cause for the most frequent defects.",
              "supporting_data": "e.g., '{top_defects[0]['DefectName']}' accounts for a significant portion of scraps."
            }}
          ],
          "recommendations": [
            {{
              "title": "A short title for the action.",
              "description": "A concrete, actionable step.",
              "priority": "High, Medium, or Low"
            }}
          ]
        }}
        """

    def _create_fail_analysis_prompt(self, statistics: Dict, period_type: str) -> str:
        top_defects = "\n".join([f"- {d['defect']}: {d['count']} times" for d in statistics.get('top_defects', [])[:5]])
        top_products = "\n".join([f"- {p['product']}: {p['count']} fails" for p in statistics.get('top_products', [])[:5]])
        worst_item = statistics['top_defects'][0]['defect'] if statistics.get('top_defects') else "N/A"

        return f"""
        You are a Continuous Improvement Analyst. Analyze production FAIL data. Your response MUST be a single, valid JSON object and in English.

        **Analysis Context:**
        - Report Period: {period_type.title()}
        - Total Fails: {statistics.get('total_fails', 'N/A')}
        - Fail Rate: {statistics.get('fail_rate', 0):.2f}%

        **Top 5 Defects:**
        {top_defects}

        **Top 5 Products with Fails:**
        {top_products}

        **Task:**
        Analyze the data and provide a response in the specified JSON format. If the fail rate or concentration of a defect is high, propose a Kaizen project.

        {{
          "executive_summary": "A concise summary of findings, the most critical issue, and the quality trend.",
          "root_causes": [{{ "problem_area": "e.g., Component Quality", "cause_description": "Detailed root cause.", "supporting_data": "Data that supports this conclusion." }}],
          "recommendations": [{{ "title": "Action title", "description": "What to do.", "priority": "High", "target_problem": "The defect/product to solve." }}],
          "kaizen_project_proposal": {{
              "project_title": "Kaizen Project to Reduce '{worst_item}' Fails",
              "problem_statement": "Data-driven description of the {worst_item} problem.",
              "goal": "A SMART goal to reduce the issue.",
              "suggested_team": ["Quality Engineer", "Process Engineer"],
              "initial_steps": ["1. Deep-dive data analysis", "2. Gemba walk on the affected line."]
          }}
        }}
        If a Kaizen project is not necessary, set "kaizen_project_proposal" to null.
        """

    def _create_stoppage_analysis_prompt(self, statistics: Dict, production_data: Dict, period_type: str) -> str:
        top_freq = "\n".join([f"- {p}: {c} times" for p, c in statistics.get('top_problems_by_freq', [])])
        top_time = "\n".join([f"- {p}: {t:.2f} hours" for p, t in statistics.get('top_problems_by_time', [])])
        worst_problem = statistics['top_problems_by_time'][0][0] if statistics.get('top_problems_by_time') else "N/A"

        # --- NUOVA SEZIONE: ISTRUZIONI SPECIFICHE PER L'ESPERTO DI PROCESSO ---
        expert_knowledge_prompt = """
        **Expert Process Knowledge:**
        You MUST incorporate the following expert knowledge into your analysis:
        - The code **"CHO"** means **"Change Over"**. This is the time spent setting up the production line for a new product/order.
        - **High CHO time is a critical issue.** It does NOT mean a machine is broken. It indicates inefficiency in the setup process.
        - **Common Root Causes for long CHO:**
            1.  **Poor Preparation:** The setup (next job's components, tools, documentation) is not prepared before the previous job finishes.
            2.  **Incomplete Kits:** The component kits from the warehouse are missing parts. This forces the line to stop and wait for a component search, which is a major source of delay.
        - **If "CHO" is a top problem, your recommendations MUST focus on:**
            - **SMED (Single-Minute Exchange of Die) principles:** Suggest preparing the changeover *before* the current job ends (external vs. internal activities).
            - **Kit Verification:** Recommend a process to verify the completeness of component kits *before* they reach the production line.
            - **Warehouse-Production Coordination:** Emphasize the need for better communication between the warehouse and production to ensure accurate and complete kits.
        """
        # --- FINE NUOVA SEZIONE ---

        return f"""
        You are a Production Maintenance and Continuous Improvement Analyst with deep expertise in electronics manufacturing.
        Your task is to analyze line stoppage data for a {period_type} report. Your response MUST be a single, valid JSON object and in English.

        {expert_knowledge_prompt}

        **Analysis Context:**
        - Report Period: {period_type.title()}
        - Total Boards Produced in Period: {production_data.get('NrBoards', 'N/A')}
        - Total Stoppages: {statistics.get('total_stoppages', 'N/A')}
        - Total Downtime: {statistics.get('total_downtime_hours', 0):.2f} hours

        **Top 5 Problems by Frequency of Stoppage:**
        {top_freq}

        **Top 5 Problems by Total Downtime (Hours):**
        {top_time}

        **Task:**
        Analyze the provided data using your expert knowledge. If "CHO" is a significant problem, your root cause analysis and recommendations must reflect the specific process issues related to Change Overs. Your response MUST be a single, valid JSON object.

        {{
          "executive_summary": "A concise summary of findings, highlighting the main cause of downtime (especially if it is 'CHO') and the overall line performance.",
          "root_causes": [{{ 
              "problem_area": "e.g., Change Over Process, Machine Failure, Material Supply", 
              "cause_description": "Detailed root cause. If the problem is 'CHO', explain it in terms of preparation or kit issues.", 
              "supporting_data": "Data that supports this conclusion (e.g., ''CHO' is the number one cause of downtime, indicating a systemic process issue')." 
          }}],
          "recommendations": [{{ 
              "title": "Action title (e.g., 'Implement Pre-Changeover Kit Verification').", 
              "description": "What to do. If targeting 'CHO', suggest specific actions like preparing kits in advance or verifying their contents.", 
              "priority": "High, Medium, or Low",
              "target_problem": "The problem to solve (e.g., 'CHO')."
          }}],
          "kaizen_project_proposal": {{
              "project_title": "Kaizen Project to Optimize '{worst_problem}' Process",
              "problem_statement": "Data-driven description of the '{worst_problem}' issue and its impact on downtime and efficiency.",
              "goal": "A SMART goal to reduce the downtime/impact of this issue (e.g., 'Reduce average CHO time by 25% within 3 months').",
              "suggested_team": ["Production Supervisor", "Warehouse Lead", "Process Engineer", "Quality Technician"],
              "initial_steps": ["1. Map the current Change Over process (Value Stream Mapping).", "2. Time and record all activities during 5 recent CHOs.", "3. Analyze the completeness of the last 20 component kits."]
          }}
        }}

        If "CHO" is NOT the main problem, you can propose a more generic Kaizen project for the worst problem. 
        If a Kaizen project is not necessary at all, set "kaizen_project_proposal" to null.
        """