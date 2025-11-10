"""
AI Analyzer per Scrap Analysis
Utilizza Ollama con Llama 3.2 per analisi root cause
"""
import json
from typing import Dict, List, Any
from logger_config import setup_logger

logger = setup_logger('AIAnalyzer')


class AIAnalyzer:
    """Analizza difetti usando AI con Ollama"""

    def __init__(self, api_key: str = None, provider: str = "ollama"):
        """
        Inizializza AI Analyzer

        Args:
            api_key: Non utilizzato per Ollama
            provider: "ollama" per uso locale
        """
        self.provider = provider
        self.model = "llama3.2"  # Modello specifico per Ollama
        self.logger = logger

        # Verifica se Ollama è disponibile
        if provider == "ollama":
            try:
                import ollama
                self.client = ollama.Client()
                # Test connessione a Ollama
                self.client.list()  # Test semplice
                self.use_ai = True
                self.logger.info("✅ Ollama inizializzato con modello llama3.2")
            except ImportError:
                self.logger.error("❌ Ollama non installato. Installare con: pip install ollama")
                self.use_ai = False
            except Exception as e:
                self.logger.warning(f"⚠️ Ollama non disponibile: {e} - Uso analisi statistica")
                self.use_ai = False
        else:
            self.use_ai = False
            self.logger.warning("⚠️ Provider AI non supportato - Uso analisi statistica")

    def analyze_defects(self, defects_data: List[Dict], production_data: Dict) -> Dict[str, Any]:
        """
        Analizza difetti e genera insights

        Args:
            defects_data: Lista difetti con conteggi
            production_data: Dati produzione

        Returns:
            dict: Root causes e raccomandazioni
        """
        if self.use_ai and self.provider == "ollama":
            return self._ollama_analysis(defects_data, production_data)
        else:
            return self._statistical_analysis(defects_data, production_data)

    def _ollama_analysis(self, defects_data: List[Dict], production_data: Dict) -> Dict[str, Any]:
        """Analisi con Ollama e Llama 3.2"""
        try:
            # Prepara prompt per Ollama
            prompt = self._create_ollama_prompt(defects_data, production_data)

            # Chiamata a Ollama
            response = self.client.generate(
                model=self.model,
                prompt=prompt,
                options={
                    'temperature': 0.3,
                    'top_p': 0.9,
                    'num_predict': 1500  # Limite tokens
                }
            )

            result_text = response['response']

            # Estrai JSON dalla risposta (Ollama potrebbe aggiungere testo extra)
            json_start = result_text.find('{')
            json_end = result_text.rfind('}') + 1

            if json_start != -1 and json_end != -1:
                json_str = result_text[json_start:json_end]
                return json.loads(json_str)
            else:
                self.logger.warning("❌ Formato JSON non trovato nella risposta Ollama")
                return self._statistical_analysis(defects_data, production_data)

        except Exception as e:
            self.logger.error(f"❌ Errore analisi Ollama: {e}")
            return self._statistical_analysis(defects_data, production_data)

    def _create_ollama_prompt(self, defects_data: List[Dict], production_data: Dict) -> str:
        """Crea prompt ottimizzato per Llama 3.2"""

        # Formatta i dati per il prompt
        defects_summary = "\n".join([
            f"- {defect['DefectName']}: {defect['Count']} occorrenze (Area: {defect.get('TopArea', 'N/A')})"
            for defect in defects_data[:10]
        ])

        total_defects = sum(d['Count'] for d in defects_data)

        prompt = f"""Sei un esperto di analisi qualità in produzione industriale per VANDEWIELE ROMANIA SRL, specializzato in Wave Soldering.

Dati di produzione:
- Ordini totali: {production_data.get('NrOrders', 0)}
- Schede prodotte: {production_data.get('NrBoards', 0):,}
- Difetti totali: {total_defects}

Top difetti rilevati:
{defects_summary}

Analizza questi dati e fornisci:
1. Root causes principali basate sui pattern dei difetti
2. Raccomandazioni specifiche per migliorare il processo

Rispondi SOLO in formato JSON valido:

{{
  "root_causes": [
    {{
      "category": "categoria (es: Thermal Profile, Flux Management, Process Parameters)",
      "cause": "causa specifica dettagliata",
      "impact": "impatto sulla qualità e produzione"
    }}
  ],
  "recommendations": [
    {{
      "title": "titolo breve della raccomandazione",
      "description": "descrizione dettagliata dell'azione da intraprendere",
      "priority": "Alta/Media/Bassa",
      "impact": "impatto atteso sul miglioramento"
    }}
  ],
  "analysis_notes": "breve sommario dell'analisi"
}}

Considera i parametri tipici del Wave Soldering:
- Lega: SAC305
- Temperatura wave: 250-260°C
- Preheating: 110-130°C
- Velocità conveyor: 0.8-1.2 m/min
- Flux: tipo ORL0/ORL1

Focus sulla risoluzione pratica dei problemi più frequenti."""

        return prompt

    def _statistical_analysis(self, defects_data: List[Dict], production_data: Dict) -> Dict[str, Any]:
        """Analisi statistica senza AI"""
        total_defects = sum(d['Count'] for d in defects_data)
        top_5 = defects_data[:5]

        # Root causes basate su pattern comuni
        root_causes = []
        recommendations = []

        for defect in top_5:
            name = defect['DefectName'].lower()
            count = defect['Count']
            percentage = (count / total_defects * 100) if total_defects > 0 else 0

            # Pattern matching per root causes
            if 'bridge' in name or 'ponte' in name:
                root_causes.append({
                    'category': 'Thermal Profile',
                    'cause': f'Ponti termici rilevati ({count} casi, {percentage:.1f}%)',
                    'impact': 'Temperatura wave o velocità conveyor non ottimali'
                })
                recommendations.append({
                    'title': 'Ottimizzazione Profilo Termico',
                    'description': 'Verificare temperatura wave (250-260°C) e velocità conveyor (0.8-1.2 m/min)',
                    'priority': 'Alta',
                    'impact': 'Riduzione ponti termici'
                })

            elif 'cold' in name or 'freddo' in name:
                root_causes.append({
                    'category': 'Temperature',
                    'cause': f'Giunti freddi ({count} casi, {percentage:.1f}%)',
                    'impact': 'Preheating insufficiente o temperatura wave bassa'
                })
                recommendations.append({
                    'title': 'Controllo Temperature',
                    'description': 'Aumentare preheating (110-130°C) e verificare temperatura wave',
                    'priority': 'Alta',
                    'impact': 'Miglioramento qualità giunti'
                })

            elif 'flux' in name:
                root_causes.append({
                    'category': 'Flux Management',
                    'cause': f'Problemi flux ({count} casi, {percentage:.1f}%)',
                    'impact': 'Applicazione flux non uniforme o flux degradato'
                })
                recommendations.append({
                    'title': 'Gestione Flux',
                    'description': 'Verificare densità flux (0.82-0.84) e sistema applicazione',
                    'priority': 'Media',
                    'impact': 'Applicazione flux più uniforme'
                })

            elif 'hole' in name or 'foro' in name:
                root_causes.append({
                    'category': 'Process Parameters',
                    'cause': f'Riempimento fori insufficiente ({count} casi, {percentage:.1f}%)',
                    'impact': 'Tempo contatto wave insufficiente o angolo PCB errato'
                })
                recommendations.append({
                    'title': 'Parametri Processo',
                    'description': 'Ridurre velocità conveyor o aumentare angolo PCB (5-7°)',
                    'priority': 'Media',
                    'impact': 'Miglior riempimento fori'
                })

        # Raccomandazioni generali
        if total_defects > 100:
            recommendations.append({
                'title': 'Audit Processo Completo',
                'description': 'Alto numero difetti richiede audit sistematico del processo wave',
                'priority': 'Critica',
                'impact': 'Identificazione problemi sistemici'
            })

        return {
            'root_causes': root_causes,
            'recommendations': recommendations,
            'analysis_type': 'statistical',
            'total_defects_analyzed': total_defects,
            'analysis_notes': f'Analisi statistica basata su {total_defects} difetti totali'
        }