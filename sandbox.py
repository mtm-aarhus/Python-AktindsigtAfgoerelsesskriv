from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
from robot_framework.process import process
from OpenOrchestrator.database.queues import QueueElement, QueueStatus
import json
from typing import Optional
from multiprocessing import freeze_support


def make_queue_element_with_payload(
    payload: dict | list,
    queue_name: str,
    reference: Optional[str] = None,
    created_by: Optional[str] = None,
    status: QueueStatus = QueueStatus.NEW, 
) -> QueueElement:
    # Validate & serialize
    data_str = json.dumps(payload, ensure_ascii=False)
    if len(data_str) > 2000:
        raise ValueError("data exceeds 2000 chars (column limit)")

    return QueueElement(
        queue_name=queue_name,
        status=status,
        data=data_str,
        reference=reference,
        created_by=created_by,
    )
def main():
    raw_json = """{
    "AnsøgerNavn": "Jonas Holm Riis",
    "AnsøgerEmail": "h-r@live.dk",
    "Afdeling": "Byrum",
    "Aktindsigtsovermappe": "3188 - Aktindsigt Vejbyggelinje",
    "SagsbehandlerEmail": "kloma@aarhus.dk",
    "DeskProID": 3188,
    "AktindsigtsDato": "2026-05-15T00:00:00Z",
    "Lovgivning": "Andet (Genererer fuld frase) "
    }"""

    payload = json.loads(raw_json)


    qe = make_queue_element_with_payload(
        payload=payload,
        queue_name="AktbobAfgørelse",
        reference="Sandbox",
        status=QueueStatus.NEW, 
    )

    orchestrator_connection = OrchestratorConnection(
            "AktbobAfgørelsesskriv",
            os.getenv("OpenOrchestratorSQL"),
            os.getenv("OpenOrchestratorKey"),
            None,
            None,
            None
        )


    process(orchestrator_connection, qe)


if __name__ == "__main__":
    freeze_support()
    main()