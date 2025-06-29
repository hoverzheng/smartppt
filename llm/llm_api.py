import requests
import json
from typing import List, Dict, Optional

class LLMApi:
    def __init__(self, api_key: str = None, base_url: str = None, model: str = "gpt-3.5-turbo"):
        """
        初始化LLM API接口
        
        Args:
            api_key: API密钥
            base_url: API基础URL，默认为OpenAI官方地址
            model: 使用的模型名称
        """
        self.api_key = api_key
        self.base_url = base_url or "https://api.openai.com/v1"
        self.model = model
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}" if api_key else ""
        }

    def generate_outline(self, topic: str, num_pages: int) -> List[Dict]:
        """
        调用LLM生成PPT大纲
        
        Args:
            topic: PPT主题
            num_pages: PPT页数
            
        Returns:
            包含每页标题、总结和要点的列表
        """
        prompt = self._create_outline_prompt(topic, num_pages)
        
        try:
            print(f"正在调用LLM API，主题: {topic}, 页数: {num_pages}")
            response = self._call_llm(prompt)
            print(f"LLM返回内容: {response[:200]}...")  # 打印前200个字符
            result = self._parse_outline_response(response, num_pages)
            print(f"解析结果: {result}")
            return result
        except Exception as e:
            print(f"LLM调用失败: {e}")
            # 返回默认内容作为fallback
            return self._generate_fallback_content(topic, num_pages)

    def _create_outline_prompt(self, topic: str, num_pages: int) -> str:
        """创建用于生成PPT大纲的提示词"""
        return f"""
请为以下主题生成一个{num_pages}页的PPT大纲，要求内容简洁明了且与主题高度相关：

主题：{topic}

要求：
1. 每页都要有明确的标题（不超过15个字）
2. 每页开头要有1-2句话的总结性介绍，用于引出后面的内容
   - 总结性介绍应该像开场白一样，自然地引出该页要讨论的内容
   - 例如：如果主题是"成都美食"，第一页的总结可以是"成都有数不清的各种美食，从街边小吃到高档餐厅，每一种都让人流连忘返"
   - 这样的介绍能够自然地引出后面的具体美食类型
3. 每页必须包含3-4个主要论点（严格控制在3-4个，不能少于3个）
4. 每个论点要有1-2个具体的事实点或数据来支持
5. 每个事实点都要有简要说明（10-20个字），解释该事实点的含义或重要性
6. 内容逻辑清晰，层次分明
7. 第一页通常是介绍页，最后一页通常是总结页
8. 每页的内容要不同，避免重复
9. 所有文字要简洁，避免冗长
10. 内容必须与主题"{topic}"高度相关，不要生成通用的"要点1、要点2"等内容
11. 必须严格按照要求生成{num_pages}页，不能多也不能少

请以JSON格式返回，格式如下：
[
    {{
        "title": "页面标题",
        "summary": "该页的总结性介绍，1-2句话自然地引出后面的内容",
        "points": [
            {{
                "main_point": "主要论点1（简洁明了且与主题相关）",
                "supporting_facts": [
                    {{
                        "fact": "支持事实1",
                        "explanation": "简要说明（10-20个字）"
                    }},
                    {{
                        "fact": "支持事实2", 
                        "explanation": "简要说明（10-20个字）"
                    }}
                ]
            }},
            {{
                "main_point": "主要论点2（简洁明了且与主题相关）",
                "supporting_facts": [
                    {{
                        "fact": "支持事实1",
                        "explanation": "简要说明（10-20个字）"
                    }},
                    {{
                        "fact": "支持事实2",
                        "explanation": "简要说明（10-20个字）"
                    }}
                ]
            }},
            {{
                "main_point": "主要论点3（简洁明了且与主题相关）",
                "supporting_facts": [
                    {{
                        "fact": "支持事实1",
                        "explanation": "简要说明（10-20个字）"
                    }},
                    {{
                        "fact": "支持事实2",
                        "explanation": "简要说明（10-20个字）"
                    }}
                ]
            }},
            {{
                "main_point": "主要论点4（简洁明了且与主题相关）",
                "supporting_facts": [
                    {{
                        "fact": "支持事实1",
                        "explanation": "简要说明（10-20个字）"
                    }},
                    {{
                        "fact": "支持事实2",
                        "explanation": "简要说明（10-20个字）"
                    }}
                ]
            }}
        ]
    }}
]

只返回JSON格式的内容，不要其他说明文字。确保严格按照要求生成{num_pages}页，内容简洁，不会超出PPT页面范围。
"""

    def _call_llm(self, prompt: str) -> str:
        """调用LLM API"""
        if not self.api_key:
            raise ValueError("API密钥未设置")
            
        data = {
            "model": self.model,
            "messages": [
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.7,
            "max_tokens": 3000
        }
        
        print(f"发送请求到: {self.base_url}")
        print(f"使用模型: {self.model}")
        
        response = requests.post(
            f"{self.base_url}/chat/completions",
            headers=self.headers,
            json=data,
            timeout=600
        )
        
        print(f"API响应状态码: {response.status_code}")
        
        if response.status_code != 200:
            print(f"API错误响应: {response.text}")
            raise Exception(f"API调用失败: {response.status_code} - {response.text}")
            
        result = response.json()
        return result["choices"][0]["message"]["content"]

    def _parse_outline_response(self, response: str, num_pages: int) -> List[Dict]:
        """解析LLM返回的大纲内容"""
        try:
            # 清理响应文本，移除可能的markdown标记
            cleaned_response = response.strip()
            if cleaned_response.startswith('```json'):
                cleaned_response = cleaned_response[7:]
            if cleaned_response.endswith('```'):
                cleaned_response = cleaned_response[:-3]
            cleaned_response = cleaned_response.strip()
            
            print(f"清理后的响应: {cleaned_response[:200]}...")
            
            # 尝试解析JSON响应
            content = json.loads(cleaned_response)
            if isinstance(content, list):
                print(f"JSON解析成功，实际页数: {len(content)}, 期望页数: {num_pages}")
                
                # 如果页数不匹配，进行调整
                if len(content) == num_pages:
                    return content
                elif len(content) > num_pages:
                    # 如果返回的页数太多，截取前num_pages页
                    print(f"返回页数过多，截取前{num_pages}页")
                    return content[:num_pages]
                else:
                    # 如果返回的页数不够，用fallback内容补充
                    print(f"返回页数不足，用fallback内容补充")
                    fallback_content = self._generate_fallback_content("补充内容", num_pages - len(content))
                    return content + fallback_content
            else:
                print(f"JSON格式不正确，期望列表，实际{type(content)}")
        except json.JSONDecodeError as e:
            print(f"JSON解析失败: {e}")
        except Exception as e:
            print(f"解析过程出错: {e}")
        
        # 如果解析失败，尝试从文本中提取内容
        print("尝试文本解析...")
        return self._extract_content_from_text(response, num_pages)

    def _extract_content_from_text(self, text: str, num_pages: int) -> List[Dict]:
        """从文本中提取PPT内容（备用方案）"""
        lines = text.split('\n')
        pages = []
        current_page = None
        current_point = None
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # 检测页面标题
            if any(keyword in line.lower() for keyword in ['第', 'page', 'slide', '页']):
                if current_page:
                    pages.append(current_page)
                current_page = {
                    "title": line, 
                    "summary": f"{line}的主要内容概述",
                    "points": []
                }
                current_point = None
            elif current_page and line.startswith(('总结', 'summary', '概述')):
                # 提取总结
                summary = line.split(':', 1)[1] if ':' in line else line
                current_page["summary"] = summary
            elif current_page and line.startswith(('-', '•', '*', '1.', '2.', '3.')):
                # 提取主要论点
                point_text = line.lstrip('-•*123456789. ')
                if point_text:
                    current_point = {
                        "main_point": point_text,
                        "supporting_facts": []
                    }
                    current_page["points"].append(current_point)
            elif current_point and line.startswith(('  ', '\t', '  -', '  •')):
                # 提取支持事实
                fact = line.lstrip(' \t-•*')
                if fact:
                    current_point["supporting_facts"].append(fact)
        
        if current_page:
            pages.append(current_page)
            
        # 如果提取的页数不够，补充默认内容
        while len(pages) < num_pages:
            pages.append({
                "title": f"第{len(pages) + 1}页",
                "summary": f"第{len(pages) + 1}页的主要内容概述",
                "points": [
                    {
                        "main_point": f"要点{len(pages) + 1}-{i+1}",
                        "supporting_facts": [
                            {
                                "fact": f"事实{len(pages) + 1}-{i+1}-1",
                                "explanation": "简要说明"
                            },
                            {
                                "fact": f"事实{len(pages) + 1}-{i+1}-2",
                                "explanation": "简要说明"
                            }
                        ]
                    }
                    for i in range(3)
                ]
            })
        
        # 如果页数太多，截取前num_pages页
        if len(pages) > num_pages:
            pages = pages[:num_pages]
            
        return pages

    def _generate_fallback_content(self, topic: str, num_pages: int) -> List[Dict]:
        """生成默认的PPT内容（当LLM调用失败时使用）"""
        print(f"使用fallback内容生成，主题: {topic}")
        pages = []
        
        # 第一页：介绍
        pages.append({
            "title": f"{topic} - 介绍",
            "summary": f"今天我们将深入探讨{topic}这个精彩话题，它包含了丰富的内容和深刻的内涵，让我们一起揭开它的神秘面纱。",
            "points": [
                {
                    "main_point": f"什么是{topic}",
                    "supporting_facts": [
                        {
                            "fact": f"{topic}的基本定义和概念",
                            "explanation": "核心概念解释"
                        },
                        {
                            "fact": f"{topic}在相关领域中的地位",
                            "explanation": "重要性和影响力"
                        }
                    ]
                },
                {
                    "main_point": f"{topic}的重要性",
                    "supporting_facts": [
                        {
                            "fact": "对行业发展的推动作用",
                            "explanation": "促进产业升级"
                        },
                        {
                            "fact": "对用户需求的满足程度",
                            "explanation": "解决实际问题"
                        }
                    ]
                },
                {
                    "main_point": f"{topic}的发展历程",
                    "supporting_facts": [
                        {
                            "fact": f"{topic}的历史背景",
                            "explanation": "发展起源"
                        },
                        {
                            "fact": f"{topic}的发展阶段",
                            "explanation": "关键里程碑"
                        }
                    ]
                },
                {
                    "main_point": "本次演讲的主要内容",
                    "supporting_facts": [
                        {
                            "fact": f"将涵盖{topic}的各个方面",
                            "explanation": "全面分析"
                        },
                        {
                            "fact": "提供详细的分析和数据支持",
                            "explanation": "数据支撑"
                        }
                    ]
                }
            ]
        })
        
        # 中间页：主要内容
        for i in range(1, num_pages - 1):
            pages.append({
                "title": f"{topic} - 第{i+1}部分",
                "summary": f"在了解了{topic}的基础知识后，现在让我们深入探讨它的第{i+1}个重要方面，这里有许多令人惊喜的发现。",
                "points": [
                    {
                        "main_point": f"{topic}的第{i+1}个核心要点",
                        "supporting_facts": [
                            {
                                "fact": f"{topic}的具体表现和特征",
                                "explanation": "关键特征"
                            },
                            {
                                "fact": f"{topic}相关数据和统计",
                                "explanation": "数据支撑"
                            }
                        ]
                    },
                    {
                        "main_point": f"{topic}的第{i+1}个关键因素",
                        "supporting_facts": [
                            {
                                "fact": f"{topic}影响因素分析",
                                "explanation": "驱动因素"
                            },
                            {
                                "fact": f"{topic}市场反应和反馈",
                                "explanation": "市场表现"
                            }
                        ]
                    },
                    {
                        "main_point": f"{topic}的第{i+1}个发展趋势",
                        "supporting_facts": [
                            {
                                "fact": f"{topic}当前发展状况",
                                "explanation": "现状分析"
                            },
                            {
                                "fact": f"{topic}未来发展方向",
                                "explanation": "前景展望"
                            }
                        ]
                    },
                    {
                        "main_point": f"{topic}的第{i+1}个应用案例",
                        "supporting_facts": [
                            {
                                "fact": f"{topic}实际应用案例",
                                "explanation": "实践案例"
                            },
                            {
                                "fact": f"{topic}成功经验分享",
                                "explanation": "经验总结"
                            }
                        ]
                    }
                ]
            })
        
        # 最后一页：总结
        if num_pages > 1:
            pages.append({
                "title": f"{topic} - 总结",
                "summary": f"经过深入的探讨，我们对{topic}有了全面的了解，现在让我们回顾一下最重要的发现和未来的发展方向。",
                "points": [
                    {
                        "main_point": f"{topic}主要内容回顾",
                        "supporting_facts": [
                            {
                                "fact": f"涵盖的{topic}关键知识点",
                                "explanation": "核心要点"
                            },
                            {
                                "fact": f"{topic}重要的数据和案例",
                                "explanation": "数据支撑"
                            }
                        ]
                    },
                    {
                        "main_point": f"{topic}关键要点总结",
                        "supporting_facts": [
                            {
                                "fact": f"{topic}最重要的几个方面",
                                "explanation": "重点内容"
                            },
                            {
                                "fact": f"{topic}需要重点关注的内容",
                                "explanation": "关注重点"
                            }
                        ]
                    },
                    {
                        "main_point": f"{topic}未来展望",
                        "supporting_facts": [
                            {
                                "fact": f"{topic}发展趋势预测",
                                "explanation": "发展前景"
                            },
                            {
                                "fact": f"{topic}潜在机遇分析",
                                "explanation": "机会识别"
                            }
                        ]
                    },
                    {
                        "main_point": f"{topic}行动建议",
                        "supporting_facts": [
                            {
                                "fact": f"{topic}具体实施建议",
                                "explanation": "操作指导"
                            },
                            {
                                "fact": f"{topic}下一步行动计划",
                                "explanation": "行动方案"
                            }
                        ]
                    }
                ]
            })
            
        return pages

    def set_config(self, api_key: str = None, base_url: str = None, model: str = None):
        """动态更新配置"""
        if api_key is not None:
            self.api_key = api_key
            self.headers["Authorization"] = f"Bearer {api_key}"
        if base_url is not None:
            self.base_url = base_url
        if model is not None:
            self.model = model

    def test_connection(self) -> bool:
        """测试API连接是否正常"""
        try:
            test_prompt = "请回复'连接测试成功'"
            response = self._call_llm(test_prompt)
            return "连接测试成功" in response
        except Exception as e:
            print(f"连接测试失败: {e}")
            return False 