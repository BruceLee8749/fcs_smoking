import pytest
from datetime import datetime

if __name__ == '__main__':
    """执行测试用例生成html报告"""
    now_time = datetime.now().strftime('%Y_%m_%d_%H_%M')
    pytest.main(['-vs', '--capture=sys','--html=FCS功能验证测试报告/%s.html' % now_time])
