class ContentPlanner:
    def __init__(self, llm_api):
        self.llm_api = llm_api

    def plan_content(self, topic: str, num_pages: int):
        """
        根据主题和页数，调用LLM规划PPT内容。
        返回格式：[{"title": "页面标题", "points": ["要点1", "要点2", ...]}, ...]
        """
        # 调用LLM API生成内容
        return self.llm_api.generate_outline(topic, num_pages) 