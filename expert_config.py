# expert_config.py - Configurazione esperto per analisi approfondite
EXPERT_CONFIG = {
    "company_context": {
        "name": "VANDEWIELE ROMANIA SRL",
        "department": "PTHM",
        "industry": "Elettronica per Macchinari Tessili",
        "quality_standards": ["IPC-A-610", "ISO-9001"]
    },

    "process_expertise": {
        "wave_soldering": {
            "typical_defects": [
                "solder_bridges", "cold_joints", "insufficient_wetting",
                "excess_solder", "voids", "component_shift", "tombstoning"
            ],
            "critical_parameters": [
                "wave_temperature", "preheat_temperature", "conveyor_speed",
                "flux_density", "solder_alloy_composition", "board_design"
            ]
        }
    },

    "analysis_depth": {
        "statistical_thresholds": {
            "high_priority": 5.0,  # % scrap rate per priorità alta
            "medium_priority": 2.0,
            "low_priority": 1.0
        },
        "include_cost_analysis": True,
        "include_preventive_measures": True,
        "include_risk_assessment": True
    },

    "reporting": {
        "language": "italian",
        "technical_level": "detailed",
        "include_visual_recommendations": True,
        "actionable_insights": True
    }
}