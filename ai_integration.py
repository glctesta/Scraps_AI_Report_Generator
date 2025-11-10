import requests
import json
import logging
from datetime import datetime

logging.basicConfig(level=logging.INFO)


class OllamaAIAnalyzer:
    """
    Classe per integrazione con Ollama AI per analisi difetti Wave Soldering
    """

    def __init__(self, base_url="http://localhost:11434", model_name="llama3.2"):
        """
        Inizializza l'analyzer Ollama

        Args:
            base_url (str): URL base del server Ollama
            model_name (str): Nome del modello da utilizzare
        """
        self.base_url = base_url
        self.model_name = model_name
        self.api_endpoint = f"{base_url}/api/generate"
        self.timeout = 60  # Timeout in secondi
        self.available = False

        # Inizializza e verifica connessione
        self._initialize_client()

    def _initialize_client(self):
        """
        Verifica la disponibilità del server Ollama e del modello
        """
        try:
            # Test connessione al server
            response = requests.get(f"{self.base_url}/api/tags", timeout=5)

            if response.status_code == 200:
                models_data = response.json()
                available_models = [model['name'] for model in models_data.get('models', [])]

                if self.model_name in available_models:
                    self.available = True
                    logging.info(f"✅ Ollama AI disponibile - Modello: {self.model_name}")
                    logging.info(f"   Server: {self.base_url}")
                    logging.info(f"   Modelli disponibili: {', '.join(available_models)}")
                else:
                    logging.warning(f"⚠️ Modello {self.model_name} non trovato")
                    logging.warning(f"   Modelli disponibili: {', '.join(available_models)}")
                    logging.warning(f"   Usa: ollama pull {self.model_name}")
                    self.available = False
            else:
                logging.warning(f"⚠️ Ollama server risponde con status {response.status_code}")
                self.available = False

        except requests.exceptions.ConnectionError:
            logging.warning(f"⚠️ Impossibile connettersi a Ollama su {self.base_url}")
            logging.warning("   Verifica che Ollama sia in esecuzione: ollama serve")
            self.available = False
        except Exception as e:
            logging.error(f"❌ Errore inizializzazione Ollama: {e}")
            self.available = False

    def _call_ollama(self, prompt, temperature=0.7, max_tokens=2000):
        """
        Effettua una chiamata al server Ollama

        Args:
            prompt (str): Prompt da inviare al modello
            temperature (float): Temperatura per la generazione (0.0-1.0)
            max_tokens (int): Numero massimo di token da generare

        Returns:
            str: Risposta del modello o None in caso di errore
        """
        if not self.available:
            logging.warning("⚠️ Ollama non disponibile, skip AI call")
            return None

        try:
            payload = {
                "model": self.model_name,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": temperature,
                    "num_predict": max_tokens
                }
            }

            logging.debug(f"🤖 Calling Ollama API: {self.api_endpoint}")
            logging.debug(f"   Model: {self.model_name}, Temperature: {temperature}")

            response = requests.post(
                self.api_endpoint,
                json=payload,
                timeout=self.timeout
            )

            if response.status_code == 200:
                result = response.json()
                generated_text = result.get('response', '')

                logging.debug(f"✅ Ollama response received ({len(generated_text)} chars)")
                return generated_text
            else:
                logging.error(f"❌ Ollama API error: {response.status_code}")
                logging.error(f"   Response: {response.text[:200]}")
                return None

        except requests.exceptions.Timeout:
            logging.error(f"❌ Timeout chiamata Ollama (>{self.timeout}s)")
            return None
        except requests.exceptions.RequestException as e:
            logging.error(f"❌ Errore richiesta Ollama: {e}")
            return None
        except Exception as e:
            logging.error(f"❌ Errore inaspettato Ollama: {e}", exc_info=True)
            return None

    def is_available(self):
        """
        Verifica se il servizio Ollama è disponibile

        Returns:
            bool: True se disponibile, False altrimenti
        """
        return self.available

    def generate_enhanced_recommendations(self, defect_data, analysis_results):
        """
        Genera raccomandazioni avanzate basate sui dati dei difetti Wave Soldering

        Args:
            defect_data (list): Lista di dizionari con dati grezzi dei difetti
            analysis_results (dict): Risultati dell'analisi statistica

        Returns:
            dict: Raccomandazioni strutturate
        """
        try:
            # ===== ESTRAZIONE DATI =====
            total_defects = analysis_results.get('total_defects', 0)
            defect_types = analysis_results.get('defect_distribution', {})
            machines = analysis_results.get('machine_distribution', {})
            trends = analysis_results.get('trends', {})
            critical_issues = analysis_results.get('critical_issues', [])

            # Identifica il difetto più comune
            most_common_defect = max(defect_types.items(), key=lambda x: x[1])[0] if defect_types else "Unknown"
            most_common_count = defect_types.get(most_common_defect, 0)

            # Identifica la macchina più problematica
            most_problematic_machine = max(machines.items(), key=lambda x: x[1])[0] if machines else "Unknown"
            machine_defect_count = machines.get(most_problematic_machine, 0)

            # Calcola percentuali
            defect_percentage = (most_common_count / total_defects * 100) if total_defects > 0 else 0
            machine_percentage = (machine_defect_count / total_defects * 100) if total_defects > 0 else 0

            # Determina trend
            trend_status = trends.get('overall_trend', 'stable')

            logging.info(f"📊 Generating recommendations for {total_defects} defects")
            logging.info(
                f"   Most common: {most_common_defect} ({most_common_count} occurrences, {defect_percentage:.1f}%)")
            logging.info(
                f"   Most problematic machine: {most_problematic_machine} ({machine_defect_count} defects, {machine_percentage:.1f}%)")

            # ===== COSTRUZIONE PROMPT =====
            prompt = f"""Sei un esperto di Wave Soldering con focus su leghe SAC (Sn-Ag-Cu) lead-free.

Analizza questi dati sui difetti di saldatura Wave e genera raccomandazioni tecniche specifiche.

DATI DIFETTI:
- Totale difetti rilevati: {total_defects}
- Difetto più comune: {most_common_defect} ({most_common_count} occorrenze, {defect_percentage:.1f}% del totale)
- Distribuzione difetti: {json.dumps(defect_types, indent=2)}
- Macchina più problematica: {most_problematic_machine} ({machine_defect_count} difetti, {machine_percentage:.1f}% del totale)
- Distribuzione macchine: {json.dumps(machines, indent=2)}
- Trend generale: {trend_status}
- Problemi critici identificati: {', '.join(critical_issues) if critical_issues else 'Nessuno'}

CONTESTO TECNICO WAVE SOLDERING:
- Lega: SAC305 (Sn96.5/Ag3.0/Cu0.5)
- Temperatura bath: 250-260°C
- Preheating: 110-130°C
- Flux: ORL0/ORL1 (low-solids)
- Delta T ottimale: 120-150°C

DIFETTI COMUNI E CAUSE:
- COLD_SOLDER/INSUFFICIENT_SOLDER: Temperatura bassa, tempo contatto insufficiente, flux esaurito
- BRIDGING: Velocità troppo bassa, altezza onda eccessiva, flux troppo attivo
- ICICLES/SOLDER_SPIKES: Angolo onda errato, velocità estrazione non ottimale
- THERMAL_DAMAGE: Delta T eccessivo, preheating insufficiente
- CONTAMINATION: Flux degradato, bath contaminato, manutenzione carente

Genera raccomandazioni in formato JSON VALIDO (solo JSON, nessun testo aggiuntivo):

{{
    "priority_actions": [
        {{
            "action": "Azione specifica da intraprendere immediatamente",
            "reason": "Motivazione tecnica basata sui dati",
            "priority": "high/medium/low",
            "estimated_impact": "Percentuale riduzione difetti stimata (es: 25-30%)",
            "target_defect": "Tipo di difetto target"
        }}
    ],
    "root_causes": [
        "Causa radice 1 identificata dai dati",
        "Causa radice 2 con spiegazione tecnica"
    ],
    "process_improvements": [
        "Miglioramento processo 1 con parametri specifici",
        "Miglioramento processo 2 con valori target"
    ],
    "preventive_measures": [
        "Misura preventiva 1 per evitare ricorrenza",
        "Misura preventiva 2 con frequenza consigliata"
    ],
    "training_needs": [
        "Area formazione 1 per operatori",
        "Area formazione 2 per tecnici"
    ],
    "equipment_checks": [
        "Controllo attrezzatura 1 con frequenza",
        "Controllo attrezzatura 2 con parametri da verificare"
    ],
    "technical_insights": [
        "Insight tecnico 1 basato su metallurgia SAC",
        "Insight tecnico 2 su chimica flux o termica"
    ]
}}

IMPORTANTE: Rispondi SOLO con JSON valido, senza markdown, senza spiegazioni aggiuntive."""

            # ===== CHIAMATA OLLAMA =====
            logging.info("🤖 Calling Ollama for AI recommendations...")
            response = self._call_ollama(prompt)

            if not response:
                logging.warning("⚠️ No response from Ollama, using fallback recommendations")
                return self._generate_fallback_recommendations(analysis_results)

            # ===== PARSING RISPOSTA =====
            try:
                # Rimuovi eventuali markdown code blocks
                response = response.strip()
                if response.startswith('```'):
                    # Rimuovi ```json o ``` all'inizio
                    response = response.split('\n', 1)[1] if '\n' in response else response[3:]
                if response.endswith('```'):
                    response = response.rsplit('\n', 1)[0] if '\n' in response else response[:-3]

                # Trova il JSON nella risposta
                json_start = response.find('{')
                json_end = response.rfind('}') + 1

                if json_start != -1 and json_end > json_start:
                    json_str = response[json_start:json_end]
                    recommendations = json.loads(json_str)

                    # Validazione struttura
                    required_keys = [
                        'priority_actions', 'root_causes', 'process_improvements',
                        'preventive_measures', 'training_needs', 'equipment_checks'
                    ]

                    for key in required_keys:
                        if key not in recommendations:
                            recommendations[key] = []

                    # Aggiungi metadati
                    recommendations['metadata'] = {
                        'generated_at': datetime.now().isoformat(),
                        'total_defects_analyzed': total_defects,
                        'primary_defect': most_common_defect,
                        'primary_machine': most_problematic_machine,
                        'trend': trend_status,
                        'ai_model': self.model_name
                    }

                    logging.info(f"✅ AI recommendations generated successfully")
                    logging.info(f"   Priority actions: {len(recommendations.get('priority_actions', []))}")
                    logging.info(f"   Root causes: {len(recommendations.get('root_causes', []))}")

                    return recommendations
                else:
                    logging.error("❌ No valid JSON found in Ollama response")
                    return self._generate_fallback_recommendations(analysis_results)

            except json.JSONDecodeError as e:
                logging.error(f"❌ JSON parsing error: {e}")
                logging.debug(f"   Response preview: {response[:200]}...")
                return self._generate_fallback_recommendations(analysis_results)

        except Exception as e:
            logging.error(f"❌ Error generating recommendations: {e}", exc_info=True)
            return self._generate_fallback_recommendations(analysis_results)

    def _generate_fallback_recommendations(self, analysis_results):
        """
        Genera raccomandazioni di fallback basate su regole quando AI non è disponibile

        Args:
            analysis_results (dict): Risultati analisi statistica

        Returns:
            dict: Raccomandazioni di base strutturate
        """
        total_defects = analysis_results.get('total_defects', 0)
        defect_types = analysis_results.get('defect_distribution', {})
        machines = analysis_results.get('machine_distribution', {})

        # Identifica problemi principali
        most_common_defect = max(defect_types.items(), key=lambda x: x[1])[0] if defect_types else "Unknown"
        most_problematic_machine = max(machines.items(), key=lambda x: x[1])[0] if machines else "Unknown"

        recommendations = {
            'priority_actions': [],
            'root_causes': [],
            'process_improvements': [],
            'preventive_measures': [],
            'training_needs': [],
            'equipment_checks': [],
            'technical_insights': []
        }

        # ===== RACCOMANDAZIONI BASATE SU REGOLE =====

        # COLD SOLDER / INSUFFICIENT SOLDER
        if 'COLD_SOLDER' in most_common_defect or 'INSUFFICIENT' in most_common_defect:
            recommendations['priority_actions'].append({
                'action': f'Verificare temperatura bath su {most_problematic_machine} (target: 255-260°C per SAC305)',
                'reason': f'Rilevati {defect_types.get(most_common_defect, 0)} casi di saldatura fredda',
                'priority': 'high',
                'estimated_impact': '30-40%',
                'target_defect': most_common_defect
            })
            recommendations['root_causes'].append('Temperatura bath insufficiente o profilo termico non ottimale')
            recommendations['equipment_checks'].append(
                'Controllo giornaliero temperatura bath con termometro calibrato')
            recommendations['process_improvements'].append(
                'Aumentare temperatura preheating a 120-130°C per ridurre Delta T')

        # BRIDGING
        if 'BRIDGING' in most_common_defect or 'BRIDGE' in most_common_defect:
            recommendations['priority_actions'].append({
                'action': f'Ridurre velocità conveyor su {most_problematic_machine} (target: 0.8-1.2 m/min)',
                'reason': f'Rilevati {defect_types.get(most_common_defect, 0)} casi di ponti di saldatura',
                'priority': 'high',
                'estimated_impact': '25-35%',
                'target_defect': most_common_defect
            })
            recommendations['root_causes'].append('Velocità conveyor troppo bassa o altezza onda eccessiva')
            recommendations['equipment_checks'].append(
                'Verifica settimanale altezza onda (target: 2/3 dello spessore PCB)')
            recommendations['process_improvements'].append('Ottimizzare densità flux (SG: 0.82-0.85 a 20°C)')

        # ICICLES / SOLDER SPIKES
        if 'ICICLE' in most_common_defect or 'SPIKE' in most_common_defect:
            recommendations['priority_actions'].append({
                'action': f'Verificare angolo onda su {most_problematic_machine} (target: 5-7° per SAC)',
                'reason': f'Rilevati {defect_types.get(most_common_defect, 0)} casi di stalattiti',
                'priority': 'medium',
                'estimated_impact': '20-30%',
                'target_defect': most_common_defect
            })
            recommendations['root_causes'].append('Angolo onda non ottimale o velocità estrazione PCB troppo rapida')
            recommendations['equipment_checks'].append('Controllo mensile geometria onda e usura nozzle')

        # THERMAL DAMAGE
        if 'THERMAL' in most_common_defect or 'DAMAGE' in most_common_defect:
            recommendations['priority_actions'].append({
                'action': f'Ridurre Delta T su {most_problematic_machine} (target: 120-150°C)',
                'reason': f'Rilevati {defect_types.get(most_common_defect, 0)} casi di danni termici',
                'priority': 'high',
                'estimated_impact': '40-50%',
                'target_defect': most_common_defect
            })
            recommendations['root_causes'].append(
                'Shock termico eccessivo: preheating insufficiente o temperatura bath troppo alta')
            recommendations['process_improvements'].append('Aumentare tempo soak preheating a 60-90 secondi')

        # CONTAMINATION
        if 'CONTAMINATION' in most_common_defect or 'CONTAMIN' in most_common_defect:
            recommendations['priority_actions'].append({
                'action': f'Sostituire flux e verificare pulizia bath su {most_problematic_machine}',
                'reason': f'Rilevati {defect_types.get(most_common_defect, 0)} casi di contaminazione',
                'priority': 'high',
                'estimated_impact': '35-45%',
                'target_defect': most_common_defect
            })
            recommendations['root_causes'].append('Flux degradato o bath contaminato (ossidi, impurità)')
            recommendations['equipment_checks'].append(
                'Controllo settimanale purezza lega (XRF) e sostituzione flux ogni 8 ore')

        # ===== RACCOMANDAZIONI GENERALI =====

        # Training
        if total_defects > 50:
            recommendations['training_needs'].extend([
                'Formazione operatori su parametri critici Wave Soldering SAC305',
                'Training tecnici su troubleshooting difetti comuni e ottimizzazione profilo termico',
                'Certificazione operatori secondo standard IPC-A-610 Classe 2/3'
            ])

        # Preventive measures
        recommendations['preventive_measures'].extend([
            'Implementare checklist giornaliera parametri macchina (temperatura, velocità, altezza onda)',
            'Monitoraggio continuo FPY (First Pass Yield) per identificare derive processo',
            'Manutenzione preventiva mensile: pulizia nozzle, verifica termocoppie, calibrazione sensori'
        ])

        # Technical insights
        recommendations['technical_insights'].extend([
            f'Analisi {total_defects} difetti: concentrazione su {most_common_defect} ({defect_types.get(most_common_defect, 0)} casi)',
            f'Macchina {most_problematic_machine} richiede attenzione prioritaria ({machines.get(most_problematic_machine, 0)} difetti)',
            'Leghe SAC richiedono controllo rigoroso Delta T per evitare shock termico componenti'
        ])

        # Metadata
        recommendations['metadata'] = {
            'generated_at': datetime.now().isoformat(),
            'total_defects_analyzed': total_defects,
            'primary_defect': most_common_defect,
            'primary_machine': most_problematic_machine,
            'source': 'rule_based_fallback',
            'ai_model': 'N/A (fallback mode)'
        }

        logging.info("✅ Fallback recommendations generated")
        logging.info(f"   Priority actions: {len(recommendations['priority_actions'])}")
        logging.info(
            f"   Total recommendations: {sum(len(v) for v in recommendations.values() if isinstance(v, list))}")

        return recommendations


# 🧪 TEST
if __name__ == "__main__":
    print("🧪 Test Ollama Integration\n")

    try:
        ai = OllamaAIAnalyzer(model_name="phi3:mini")

        # Test 1: Analisi testo
        print("=" * 60)
        print("TEST 1: Analisi Testo")
        print("=" * 60)
        test_text = "Replace flux nozzles on Wave machine #2 with model XYZ within 3 days"
        print(f"📝 Input: {test_text}\n")

        result = ai.analyze_text(test_text)
        print("📊 Output:")
        print(json.dumps(result, indent=2, ensure_ascii=False))

        # Test 2: Raccomandazioni
        print("\n" + "=" * 60)
        print("TEST 2: Raccomandazioni Difetti")
        print("=" * 60)

        defect_data = {
            'total_defects': 45,
            'defect_types': {
                'bridging': 18,
                'insufficient_solder': 12,
                'cold_joint': 10,
                'solder_balls': 5
            },
            'machines': {
                'Wave #1': 25,
                'Wave #2': 20
            },
            'trend': 'increasing',
            'critical_issues': ['bridging on Wave #1', 'temperature drift']
        }

        print(f"📊 Dati difetti:")
        print(json.dumps(defect_data, indent=2))
        print()

        recommendations = ai.generate_enhanced_recommendations(defect_data)
        print("💡 Raccomandazioni:")
        print(json.dumps(recommendations, indent=2, ensure_ascii=False))

        print("\n✅ Test completati con successo!")

    except Exception as e:
        logging.error(f"❌ Test failed: {e}")
        import traceback

        traceback.print_exc()
