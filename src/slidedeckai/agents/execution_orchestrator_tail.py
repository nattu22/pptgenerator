def _batch_validate_placeholder_roles(self, section, placeholder_map: dict) -> dict:
    """
    Batch-validate placeholder roles with a single LLM call.
    Returns a mapping {ph_id: role}
    """
    # This function is intended to be copied into ExecutionOrchestrator class body.
    roles = ['subtitle', 'chart', 'table', 'kpi', 'content', 'main_content', 'image', 'icon']
    items = []
    for pid, info in placeholder_map.items():
        try:
            pid_int = int(pid)
        except Exception:
            pid_int = pid
        items.append({
            'id': pid_int,
            'type': info.get('type'),
            'area': round(float(info.get('area', 0)), 2),
            'bbox': info.get('bbox'),
            'inferred_role': info.get('role')
        })

    prompt = {
        'section_title': getattr(section, 'section_title', ''),
        'section_purpose': getattr(section, 'section_purpose', ''),
        'placeholders': items,
        'allowed_roles': roles
    }

    instruction = (
        "Given the section context and the list of placeholders, return a JSON object mapping"
        " placeholder ids to the single best role from the allowed_roles list. Return only JSON."
    )

    messages = [
        {"role": "system", "content": "You are a concise classifier. Return only valid JSON mapping ids to one role."},
        {"role": "user", "content": instruction + "\n\n" + json.dumps(prompt)}
    ]

    try:
        resp = self.content_generator.client.chat.completions.create(
            model=self.content_generator.model,
            messages=messages,
            temperature=0.0,
            max_tokens=400
        )

        text = resp.choices[0].message.content.strip()
        try:
            parsed = json.loads(text)
        except Exception:
            import re
            m = re.search(r"\{[\s\S]*\}", text)
            if m:
                parsed = json.loads(m.group(0))
            else:
                raise

        result = {}
        for k, v in parsed.items():
            try:
                idx = int(k)
            except Exception:
                try:
                    idx = int(float(k))
                except Exception:
                    continue
            role = str(v).lower()
            if role not in roles:
                for r in roles:
                    if r in role:
                        role = r
                        break
                else:
                    role = placeholder_map.get(idx, {}).get('role')
            result[idx] = role

        return result
    except Exception as e:
        logger.debug(f"Batch role validation failed: {e}")
        return {int(pid): info.get('role') for pid, info in placeholder_map.items()}
