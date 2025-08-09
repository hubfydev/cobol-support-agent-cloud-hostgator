import requests
import json


class OllamaClient:
    def __init__(self, host: str, model: str):
        self.host = host.rstrip("/")
        self.model = model

    def generate_json(self, system_prompt: str, user_prompt: str, timeout=60):
        """
        Tenta gerar um JSON via /api/generate estruturado.
        Requer que o modelo suporte formatações tipo JSON.
        """
        url = f"{self.host}/api/generate"
        prompt = (
            f"{system_prompt.strip()}\n\n"
            "Responda apenas em JSON válido. Agora a tarefa do usuário:\n"
            f"{user_prompt.strip()}"
        )
        payload = {
            "model": self.model,
            "prompt": prompt,
            "stream": False,
            "options": {"temperature": 0.2}
        }
        r = requests.post(url, json=payload, timeout=timeout)
        r.raise_for_status()
        data = r.json()
        txt = data.get("response", "").strip()
        # tenta JSON parse
        try:
            return json.loads(txt)
        except Exception:
            # heurística: extrair primeiro bloco { ... }
            import re
            m = re.search(r"\{.*\}", txt, re.DOTALL)
            if m:
                return json.loads(m.group(0))
            raise
