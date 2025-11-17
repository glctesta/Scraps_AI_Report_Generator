"""
Fail Analyzer - Analisi fail di produzione mensili e settimanali
"""
import json
from datetime import datetime, timedelta
from typing import Dict, List, Any
from logger_config import setup_logger

logger = setup_logger('FailAnalyzer')


class FailAnalyzer:
    """Analizza i fail di produzione usando AI"""

    def __init__(self, ai_analyzer):
        """
        Inizializza Fail Analyzer

        Args:
            ai_analyzer: Istanza di AIAnalyzer per analisi AI
        """
        self.ai_analyzer = ai_analyzer
        self.logger = logger

    def get_fail_data(self, db_connection, start_date: str, end_date: str) -> List[Dict]:
        """
        Recupera dati FAILS (non scraps) dal database usando la query corretta
        """
        try:
            query = """                    
                    SELECT A.ProductCode, 
                           A.OrderProduction, 
                           A.DataVerify, 
                           A.IDBoard, 
                           A.Defects, 
                           A.Riferiments, 
                           A.PhaseName as Area, 
                           A.Name      as Operator
                    FROM (SELECT Products.ProductCode, 
                                 Orders.OrderProduction, 
                                 OrderPhases.PhasePosition, 
                                 Phases.PhaseName, 
                                 Users.Name, 
                                 QualityVerify.DataVerify, 
                                 QualityVerify.IDBoard, 
                                 Traceability_RS.dbo.BoardLabels(Boards.IDBoard) AS Labels, 
                                 null                            as IDBox, 
                                 ''                              as BoxCode, 
                                 Defects.DefectNameRO            as Defects, 
                                 Riferiments.CodRiferimento      as Riferiments 
                          FROM traceability_rs.dbo.QualityVerify 
                                   INNER JOIN traceability_rs.dbo.OrderPhases 
                                              ON QualityVerify.IDOrderPhase = OrderPhases.IDOrderPhase 
                                   INNER JOIN traceability_rs.dbo.Phases ON OrderPhases.IDPhase = Phases.IDPhase 
                                   INNER JOIN traceability_rs.dbo.Users ON QualityVerify.IDUser = Users.IDUser 
                                   INNER JOIN traceability_rs.dbo.Boards ON QualityVerify.IDBoard = Boards.IDBoard 
                                   INNER JOIN traceability_rs.dbo.Orders ON OrderPhases.IDOrder = Orders.IDOrder 
                                   INNER JOIN traceability_rs.dbo.Products ON Orders.IDProduct = Products.IDProduct 
                                   INNER JOIN traceability_rs.dbo.QualityVerifyDefects 
                                              ON QualityVerifyDefects.IDQualityVerify = QualityVerify.IDQualityVerify 
                                   INNER JOIN traceability_rs.dbo.Defects 
                                              ON Defects.IdDefect = QualityVerifyDefects.IDDefect 
                                   INNER JOIN traceability_rs.dbo.QualityVerifyDefectsRiferiments 
                                              ON QualityVerifyDefects.IDQualityVerifyDefects = 
                                                 QualityVerifyDefectsRiferiments.IDQualityVerifyDefects 
                                   INNER JOIN traceability_rs.dbo.Riferiments 
                                              ON QualityVerifyDefectsRiferiments.IDDibaRiferimento = 
                                                 Riferiments.IDDibaRiferimento 
                          WHERE QualityVerify.IsPass = 0 
                            AND CAST(QualityVerify.DataVerify AS DATE) BETWEEN ? AND ? 

                          UNION 

                          SELECT Products.ProductCode, 
                                 Orders.OrderProduction, 
                                 '990'                           as PhasePosition, 
                                 'Packing'                       as PhaseName, 
                                 Users.Name, 
                                 QualityVerifyBoxes.DataVerify, 
                                 Boards.IDBoard, 
                                 Traceability_RS.dbo.BoardLabels(Boards.IDBoard) AS Labels, 
                                 QualityVerifyBoxes.IDBox, 
                                 Boxes.BoxCode, 
                                 QualityVerifyBoxBoards.Reason   as Defects, 
                                 ''                              as Riferiments 
                          FROM traceability_rs.dbo.QualityVerifyBoxes 
                                   INNER JOIN traceability_rs.dbo.QualityVerifyBoxBoards 
                                              ON QualityVerifyBoxes.IDQualityVerifyBox = 
                                                 QualityVerifyBoxBoards.IDQualityVerifyBox 
                                   INNER JOIN traceability_rs.dbo.Boards 
                                              ON QualityVerifyBoxBoards.IDBoard = Boards.IDBoard 
                                   INNER JOIN traceability_rs.dbo.Orders ON Boards.IDOrder = Orders.IDOrder 
                                   INNER JOIN traceability_rs.dbo.Boxes ON Boxes.IdBox = QualityVerifyBoxes.IdBox 
                                   INNER JOIN traceability_rs.dbo.Users ON QualityVerifyBoxes.IDUser = Users.IDUser 
                                   INNER JOIN traceability_rs.dbo.Products ON Orders.IDProduct = Products.IDProduct 
                          WHERE QualityVerifyBoxBoards.IsPass = 0 
                          AND CAST(QualityVerifyBoxes.DataVerify AS DATE) BETWEEN ? AND ?) A
                    ORDER BY A.ProductCode, 
                             A.OrderProduction, 
                             A.PhasePosition 
                    """

            cursor = db_connection.cursor()

            # Converti le date per il formato SQL Server
            start_date_sql = f"{start_date} 00:00:00"
            end_date_sql = f"{end_date} 23:59:59"

            cursor.execute(query, (start_date_sql, end_date_sql, start_date_sql, end_date_sql))

            fails = []
            for row in cursor.fetchall():
                fails.append({
                    'FailID': getattr(row, 'IDBoard', None),  # Usa IDBoard come identificativo
                    'FailDate': getattr(row, 'DataVerify', None),
                    'ProductCode': getattr(row, 'ProductCode', 'Unknown'),
                    'DefectType': getattr(row, 'Defects', 'Unknown'),
                    'Area': getattr(row, 'Area', 'Unknown'),
                    'Operator': getattr(row, 'Operator', 'Unknown'),
                    'BoardID': getattr(row, 'IDBoard', None),
                    'OrderProduction': getattr(row, 'OrderProduction', 'Unknown'),
                    'Riferiments': getattr(row, 'Riferiments', ''),
                    'Count': 1  # Ogni riga è un fail
                })

            cursor.close()
            logger.info(f"Recuperati {len(fails)} record FAIL (non scraps)")

            # Log di debug per i primi 3 record
            if fails:
                for i, fail in enumerate(fails[:3]):
                    logger.debug(f"Fail sample {i + 1}: {fail}")

            return fails

        except Exception as e:
            logger.error(f"Errore recupero dati FAIL: {e}", exc_info=True)
            return []

    def _calculate_fail_statistics(self, fail_data: List[Dict], production_data: Dict) -> Dict:
        """Calcola statistiche per i FAILS"""
        try:
            total_fails = len(fail_data)
            total_boards = production_data.get('NrBoards', 1)

            # Calcola fail rate
            fail_rate = (total_fails / total_boards * 100) if total_boards > 0 else 0

            # Statistiche per tipo di difetto - USA SEMPRE 1 COME COUNT
            defect_stats = {}
            for fail in fail_data:
                defect_type = fail.get('DefectType', 'Unknown')
                if defect_type not in defect_stats:
                    defect_stats[defect_type] = 0
                defect_stats[defect_type] += 1  # Ogni fail conta come 1

            # Top defects - CORREGGI QUI: usa 'count' invece di 'Count'
            top_defects = sorted(
                [{'defect': k, 'count': v} for k, v in defect_stats.items()],  # 'count' minuscolo
                key=lambda x: x['count'],
                reverse=True
            )

            # Statistiche per prodotto
            product_stats = {}
            for fail in fail_data:
                product = fail.get('ProductCode', 'Unknown')
                if product not in product_stats:
                    product_stats[product] = 0
                product_stats[product] += 1

            top_products = sorted(
                [{'product': k, 'count': v} for k, v in product_stats.items()],  # 'count' minuscolo
                key=lambda x: x['count'],
                reverse=True
            )

            # Statistiche per area/fase
            area_stats = {}
            for fail in fail_data:
                area = fail.get('Area', 'Unknown')
                if area not in area_stats:
                    area_stats[area] = 0
                area_stats[area] += 1

            top_areas = sorted(
                [{'area': k, 'count': v} for k, v in area_stats.items()],  # 'count' minuscolo
                key=lambda x: x['count'],
                reverse=True
            )

            logger.info(
                f"Statistiche FAIL: {total_fails} fails, {fail_rate:.2f}% rate, {len(defect_stats)} tipi difetti")

            return {
                'total_fails': total_fails,
                'total_boards': total_boards,
                'fail_rate': fail_rate,
                'defect_stats': defect_stats,
                'top_defects': top_defects[:10],
                'top_products': top_products[:10],
                'top_areas': top_areas[:5],
                'unique_defects': len(defect_stats),
                'unique_products': len(product_stats),
                'unique_areas': len(area_stats)
            }

        except Exception as e:
            logger.error(f"Errore calcolo statistiche FAIL: {e}")
            return self._calculate_basic_statistics(fail_data)

    def _calculate_basic_statistics(self, fail_data: List[Dict]) -> Dict:
        """Calcola statistiche di base quando l'analisi dettagliata fallisce"""
        total_fails = len(fail_data)

        return {
            'total_fails': total_fails,
            'total_boards': total_fails * 10,  # Stima
            'fail_rate': 0.0,
            'defect_stats': {},
            'top_defects': [],
            'top_products': [],
            'unique_defects': 0,
            'unique_products': 0
        }

    def _calculate_empty_statistics(self) -> Dict:
        """Statistiche quando non ci sono dati"""
        return {
            'total_fails': 0,
            'total_boards': 0,
            'fail_rate': 0.0,
            'defect_stats': {},
            'top_defects': [],
            'top_products': [],
            'unique_defects': 0,
            'unique_products': 0
        }

    def calculate_fail_statistics(self, fail_data: List[Dict], production_data: Dict) -> Dict[str, Any]:
        """
        Calcola statistiche avanzate sui fail

        Args:
            fail_data: Lista di fail
            production_data: Dati produzione

        Returns:
            Dizionario con statistiche
        """
        total_fails = len(fail_data)
        total_boards = production_data.get('NrBoards', 0)

        # Calcola fail rate
        fail_rate = (total_fails / total_boards * 100) if total_boards > 0 else 0

        # Analisi per prodotto
        product_stats = {}
        for fail in fail_data:
            product = fail['ProductCode']
            if product not in product_stats:
                product_stats[product] = {'count': 0, 'defects': set(), 'operators': set()}
            product_stats[product]['count'] += 1
            product_stats[product]['defects'].add(fail['Defect'])
            product_stats[product]['operators'].add(fail['Operator'])

        # Top prodotti problematici
        top_products = sorted(
            [{'product': k, 'count': v['count'], 'unique_defects': len(v['defects'])}
             for k, v in product_stats.items()],
            key=lambda x: x['count'],
            reverse=True
        )[:10]

        # Analisi per tipo di difetto
        defect_stats = {}
        for fail in fail_data:
            defect = fail['Defect']
            if defect not in defect_stats:
                defect_stats[defect] = {'count': 0, 'products': set()}
            defect_stats[defect]['count'] += 1
            defect_stats[defect]['products'].add(fail['ProductCode'])

        # Top difetti
        top_defects = sorted(
            [{'defect': k, 'count': v['count'], 'affected_products': len(v['products'])}
             for k, v in defect_stats.items()],
            key=lambda x: x['count'],
            reverse=True
        )[:10]

        # Analisi per operatore
        operator_stats = {}
        for fail in fail_data:
            operator = fail['Operator']
            if operator not in operator_stats:
                operator_stats[operator] = {'count': 0, 'defects': set()}
            operator_stats[operator]['count'] += 1
            operator_stats[operator]['defects'].add(fail['Defect'])

        # Trend temporale (semplificato)
        date_stats = {}
        for fail in fail_data:
            date_str = fail['DataVerify'].strftime('%Y-%m-%d') if hasattr(fail['DataVerify'], 'strftime') else str(
                fail['DataVerify'])[:10]
            if date_str not in date_stats:
                date_stats[date_str] = 0
            date_stats[date_str] += 1

        return {
            'total_fails': total_fails,
            'total_boards': total_boards,
            'fail_rate': fail_rate,
            'product_stats': product_stats,
            'defect_stats': defect_stats,
            'operator_stats': operator_stats,
            'date_stats': date_stats,
            'top_products': top_products,
            'top_defects': top_defects
        }

    def analyze_fails(self, fail_data: List[Dict], production_data: Dict, period_type: str) -> Dict:
        """
        Analizza i fail con gestione errori migliorata
        """
        try:
            if not fail_data:
                return {
                    'statistics': self._calculate_empty_statistics(),
                    'ai_insights': {'analysis_type': 'no_data'},
                    'period_type': period_type,
                    'success': False
                }

            # Calcola statistiche
            statistics = self._calculate_fail_statistics(fail_data, production_data)

            # Analisi AI
            ai_insights = self.ai_analyzer.analyze_fails(fail_data, statistics, period_type)

            return {
                'statistics': statistics,
                'ai_insights': ai_insights,
                'period_type': period_type,
                'success': True
            }

        except Exception as e:
            logger.error(f"Errore analisi fail: {e}")
            return {
                'statistics': self._calculate_basic_statistics(fail_data),
                'ai_insights': {'analysis_type': 'error', 'error': str(e)},
                'period_type': period_type,
                'success': False
            }

    def _ai_fail_analysis(self, analysis_data: Dict) -> Dict[str, Any]:
        """
        Analisi AI dei fail
        """
        statistics = analysis_data['statistics']
        period_type = analysis_data['period_type']
        top_defects = analysis_data['top_defects']
        top_products = analysis_data['top_products']

        # Prepara prompt specifico per fail analysis
        prompt = self._create_fail_analysis_prompt(statistics, period_type, top_defects, top_products)

        # Usa l'AI analyzer esistente
        return self.ai_analyzer.analyze_defects(top_defects, {'NrBoards': statistics['total_boards']})

    def _create_fail_analysis_prompt(self, statistics: Dict, period_type: str, top_defects: List,
                                     top_products: List) -> str:
        """
        Crea prompt per analisi fail
        """
        total_fails = statistics['total_fails']
        total_boards = statistics['total_boards']
        fail_rate = statistics['fail_rate']

        defects_summary = "\n".join([
            f"- {defect['defect']}: {defect['count']} occorrenze (Prodotti coinvolti: {defect['affected_products']})"
            for defect in top_defects
        ])

        products_summary = "\n".join([
            f"- {product['product']}: {product['count']} fail (Tipi difetto: {product['unique_defects']})"
            for product in top_products
        ])

        prompt = f"""ANALISI FAIL DI PRODUZIONE - VANDEWIELE ROMANIA SRL

CONTESTO:
- Tipo analisi: {period_type.upper()}
- Fail totali: {total_fails:,}
- Schede controllate: {total_boards:,}
- Tasso di fail: {fail_rate:.2f}%

TOP DIFETTI RILEVATI:
{defects_summary}

TOP PRODOTTI PROBLEMATICI:
{products_summary}

RICHIESTA DI ANALISI:

Identifica le cause principali dei fail e fornisci raccomandazioni specifiche per:
1. Riduzione difetti ricorrenti
2. Miglioramento processo di controllo qualità
3. Formazione operatori
4. Ottimizzazione parametri processo

RISpondi in formato JSON:

{{
  "root_causes": [
    {{
      "category": "categoria",
      "cause": "causa specifica",
      "impact": "impatto sulla qualità"
    }}
  ],
  "recommendations": [
    {{
      "title": "titolo",
      "description": "descrizione dettagliata",
      "priority": "Alta/Media/Bassa",
      "target": "difetto/prodotto/processo target"
    }}
  ],
  "preventive_measures": [
    "misura preventiva 1",
    "misura preventiva 2"
  ]
}}"""

        return prompt