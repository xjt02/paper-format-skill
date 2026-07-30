"""paper-format 论文格式化 Skill"""
from setuptools import setup, find_packages

setup(
    name='paper-format-skill',
    version='1.0.0',
    description='将纯文本论文草稿一键排版为期刊标准格式 Word 文档（支持等宽双栏）',
    long_description=open('README.md', encoding='utf-8').read(),  # noqa: F403
    long_description_content_type='text/markdown',
    author='XJT02',
    license='MIT',
    python_requires='>=3.8',
    install_requires=[
        'python-docx>=0.8.10',
    ],
    packages=find_packages(),
    py_modules=['paper_format'],
    # 仅供 Claude Code Skill 使用时注册触发词（实际由 CLAUDE.md 中的 skill.yaml 定义）
    keywords='paper-format academic-writing word docx double-column claude skill',
    url='https://github.com/XJT02/paper-format-skill',
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Science/Research',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Topic :: Text Processing :: Markup :: Word',
    ],
)
