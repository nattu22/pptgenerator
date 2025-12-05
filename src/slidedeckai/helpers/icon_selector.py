import os
import numpy as np
import logging
from typing import Optional
from sklearn.metrics.pairwise import cosine_similarity

logger = logging.getLogger(__name__)

from slidedeckai.global_config import GlobalConfig

class IconSelector:
    def __init__(self, embeddings_path: Optional[str] = None,
                 icons_path: Optional[str] = None):
        if embeddings_path is None:
            embeddings_path = str(GlobalConfig.EMBEDDINGS_FILE_NAME)
        if icons_path is None:
            icons_path = str(GlobalConfig.ICONS_FILE_NAME)

        self.embeddings = None
        self.icons = None
        self.load_embeddings(embeddings_path, icons_path)

    def load_embeddings(self, emb_path, icons_path):
        try:
            if os.path.exists(emb_path) and os.path.exists(icons_path):
                self.embeddings = np.load(emb_path)
                self.icons = np.load(icons_path)
                logger.info(f"Loaded {len(self.icons)} icon embeddings.")
            else:
                logger.warning("Icon embeddings not found. Icon selection will be disabled.")
        except Exception as e:
            logger.error(f"Failed to load icon embeddings: {e}")

    def get_closest_icon(self, query_embedding: np.ndarray) -> Optional[str]:
        if self.embeddings is None:
            return None

        # Ensure query is 2D
        if query_embedding.ndim == 1:
            query_embedding = query_embedding.reshape(1, -1)

        similarities = cosine_similarity(query_embedding, self.embeddings)
        best_idx = np.argmax(similarities)

        return self.icons[best_idx]

    def select_icon_for_keyword(self, keyword: str, client, model=None) -> str:
        """
        Get icon filename for a keyword using embeddings.
        Fallback to 'default_icon.png' or similar if not found/error.
        """
        from slidedeckai.global_config import GlobalConfig
        if not model:
            model = GlobalConfig.LLM_EMBEDDING_MODEL

        if self.embeddings is None:
            return "placeholder.png"

        try:
            response = client.embeddings.create(
                input=keyword,
                model=model
            )
            embedding = np.array(response.data[0].embedding)
            icon_name = self.get_closest_icon(embedding)
            return icon_name if icon_name else "placeholder.png"
        except Exception as e:
            logger.error(f"Icon selection failed for '{keyword}': {e}")
            return "placeholder.png"
