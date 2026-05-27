from __future__ import annotations

class AuditNodes:
    RELEASE_FRONT = "需审阅发版工具发版账号"
    RELEASE_BACK = "需审阅发版工具后台账号"
    APP_ADMIN = "需审阅应用层管理员账号"
    APP_SCREENSHOT = "需上传应用层账号截图"
    APP_SERVER = "需审阅应用层服务器账号"
    DB_SERVER = "需审阅数据库服务器账号"
    DB_ACCOUNT = "需审阅数据库账号"
    DEV_LIST = "需上传开发人员清单"
    DEV_NO_PROD_ABOVE_READ = "开发人员在生产环境无只读以上权限"
    SOD = "无前后台管理员SOD问题"

    ALL = [
        RELEASE_FRONT,
        RELEASE_BACK,
        APP_ADMIN,
        APP_SCREENSHOT,
        APP_SERVER,
        DB_SERVER,
        DB_ACCOUNT,
        DEV_LIST,
        DEV_NO_PROD_ABOVE_READ,
        SOD,
    ]
