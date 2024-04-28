/**
 * @desc 从招商银行的邮件中提取交易信息
 */
function AccountAllEmails() {

    Logger.log('🕙 Start to Double Check All Emails From CMBC...');

    const labelFinished = GmailApp.getUserLabelByName("Accounting/Finished");
    const labelErrors = GmailApp.getUserLabelByName("Accounting/Errors");
    const labelInternal = GmailApp.getUserLabelByName("Accounting/Internal");

    if (!labelFinished || !labelErrors || !labelInternal) {
        Logger.log('ERROR: Required labels do not exist. Please check the label names.');
        return;
    }


    const threads = GmailApp.search(all_cmbc_email);


    for (let thread of threads) {
        let messages = thread.getMessages();
        for (let message of messages) {
            try {
                let transaction = analyzeTransaction(message.getPlainBody());

                switch (transaction.label) {
                    case 'Internal':
                        labelInternal.addToThread(thread);
                        message.markRead();
                        Logger.log('🔒 Transaction is internal, marked as such.');
                        break;
                    case 'Errors':
                        labelErrors.addToThread(thread);
                        message.markUnread();
                        Logger.log('❌ Transaction data is invalid, marked as error to review.');
                        break;
                    case 'Finished':
                        let payload = constructPayload(message, transaction);
                        if (createAccountingRecord(payload)) {
                            labelFinished.addToThread(thread);
                            message.markRead();
                            Logger.log('✅ Transaction has been processed successfully. ' + transaction.email);
                        } else {
                            message.markUnread();
                            labelErrors.addToThread(thread);
                            Logger.log('❌ Transaction has failed to process, but marked as unread to review.');
                        }
                        break;
                    default:
                        Logger.log('🔍 Transaction data: ' + JSON.stringify(transaction));
                        break;
                }
            } catch (error) {
                Logger.log(error.toString());
                message.markUnread(); // Ensure it's still flagged for review
                labelErrors.addToThread(thread);
            }
        }
    }
}




function analyzeTransaction(emailContent) {

    let transaction = {
        email: emailContent,
        type: '',
        catalog: '',
        amount: 0,
        label: 'Errors' // 默认标签为 Errors，会在匹配成功时修改
    };

    const patterns = [
        // 支出类别
        { pattern: /快捷支付(\d+\.?\d{0,2})元/, type: '支出', catalog: '日常支出', label: 'Finished' },
        { pattern: /银联扣款人民币(\d+\.?\d{0,2})元/, type: '支出', catalog: '转账', label: 'Finished' },
        { pattern: /信用卡还款交易人民币(\d+\.?\d{0,2})/, type: '支出', catalog: '信用卡还款', label: 'Finished' },
        { pattern: /实时转至他行人民币(\d+\.?\d{0,2})/, type: '支出', catalog: '转账', label: 'Finished' },
        { pattern: /在中铁网络一网通支付人民币(\d+\.?\d{0,2})元/, type: '支出', catalog: '交通通勤', label: 'Finished' },
        { pattern: /支付人民币(\d+\.?\d{0,2})元/, type: '支出', catalog: '日常支出', label: 'Finished' },

        // 收入类别
        { pattern: /收款(\d+\.?\d{0,2})元/, type: '收入', catalog: '其他收入', label: 'Finished' },
        { pattern: /入账工资，人民币(\d+\.?\d{0,2})，余额人民币\d+\.?\d{0,2}/, type: '收入', catalog: '工资', label: 'Finished' },
        { pattern: /他行实时转入人民币(\d+\.?\d{0,2})/, type: '收入', catalog: '转账', label: 'Finished' },
        { pattern: /入账款项，人民币(\d+\.?\d{0,2})/, type: '收入', catalog: '转账', label: 'Finished' },
        { pattern: /在中铁网络发生一网通支付退货入账人民币(\d+\.?\d{0,2})元/, type: '收入', catalog: '转账', label: 'Finished' },
        { pattern: /银联入账人民币(\d+\.?\d{0,2})元/, type: '收入', catalog: '转账', label: 'Finished' },
        { pattern: /【财付通-微信红包】退款(\d+\.?\d{0,2})元/, type: '收入', catalog: '其他收入', label: 'Finished' },
        { pattern: /银联退款人民币(\d+\.?\d{0,2})元/, type: '收入', catalog: '其他收入', label: 'Finished' },
        { pattern: /入账.*?人民币(\d+\.?\d{0,2})，/, type: '收入', catalog: '转账', label: 'Finished' },
        { pattern: /退款(\d+\.?\d{0,2})元/, type: '收入', catalog: '转账', label: 'Finished' },

        // 银行内部转账
        { pattern: /朝朝宝/, label: 'Internal' },
        { pattern: /日日宝/, label: 'Internal' },
        { pattern: /日日金/, label: 'Internal' }
    ];

    
    for (let { pattern, type, catalog, label } of patterns) {
        let match = emailContent.match(pattern);
        if (match) {
            transaction.type = type;
            transaction.catalog = catalog;
            transaction.amount = parseFloat(match[1]);
            transaction.label = label; // Assign the finished label
            return transaction;
        }
    }

    // If no patterns match, the transaction remains labeled as 'Errors'
    return transaction;
}


function constructPayload(message, transaction) {
    return {
        "编号": message.getId(),
        "分类": transaction.catalog || "其他",
        "物品或备注": message.getPlainBody(),
        "金额": transaction.amount,
        "收支": transaction.type,
        "记账时间": dayjs(message.getDate()).valueOf(),
        "邮件编码": message.getId()
    };
}


function createAccountingRecord(payload, message) {
    let emailCode = payload['邮件编码'];
    let searchResult = searchRecords(emailCode);
    if (searchResult && searchResult.data && searchResult.data.total > 0) {
        Logger.log('Record already exists, skipping...');
        return true;
    }

    let recordCreationResult = createRecord(payload);
    if (recordCreationResult && recordCreationResult.code == 0) {
        return true;
    } else {
        Logger.log('ERROR: Failed to create record, marking email for review');
        labelErrors.addToThread(message.getThread()); // 标记为错误
        return false; // 记录创建失败
    }
}


function fetchTenantAccessToken() {
    let scriptProperties = PropertiesService.getScriptProperties();
    let tokenInfo = scriptProperties.getProperties();
    let currentTime = new Date().getTime();

    if (tokenInfo.tenant_access_token && tokenInfo.expiry_time && currentTime < tokenInfo.expiry_time) {
        return {
            tenant_access_token: tokenInfo.tenant_access_token,
            expire: (tokenInfo.expiry_time - currentTime) / 1000
        };
    }

    let url = 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal';
    let payload = { app_id, app_secret };

    let options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
    };

    try {
        let response = UrlFetchApp.fetch(url, options);
        let data = JSON.parse(response.getContentText());
        if (data.code == 0 && data.tenant_access_token) {
            let expiryTime = currentTime + data.expire * 1000 - 10000; // 提前10秒过期
            scriptProperties.setProperties({
                tenant_access_token: data.tenant_access_token,
                expiry_time: expiryTime.toString()
            }, true);
            Logger.log('New token fetched and stored');
            return {
                tenant_access_token: data.tenant_access_token,
                expire: data.expire
            };
        } else {
            throw new Error('Failed to retrieve valid token');
        }
    } catch (error) {
        Logger.log('ERROR: Error fetching tenant access token: ' + error.toString());
        throw new Error('ERROR: Failed to fetch or refresh token.'); // 抛出异常确保不会标记为已读
    }
}


function searchRecords(emailCode) {
    const tenant_access_token = fetchTenantAccessToken().tenant_access_token;
    let url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${bitable_app_token}/tables/${bitable_table_id}/records/search`;
    let payload = {
        "automatic_fields": false,
        "field_names": ["邮件编码", "记账时间"],
        "filter": {
            "conditions": [{
                "field_name": "邮件编码",
                "operator": "is",
                "value": [emailCode]
            }],
            "conjunction": "and"
        },
        "sort": [{
            "desc": true,
            "field_name": "记账时间"
        }],
        "view_id": bitable_view_id
    };

    let options = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
            'Authorization': 'Bearer ' + tenant_access_token
        },
        'payload': JSON.stringify(payload)
    };

    try {
        let response = UrlFetchApp.fetch(url, options);
        let data = JSON.parse(response.getContentText());
        return data;
    } catch (error) {
        Logger.log('ERROR: Error in fetching data: ' + error.toString());
        throw new Error('Failed to search records.'); // 抛出异常确保不会标记为已读
    }
}


function createRecord(payload) {
    const tenant_access_token = fetchTenantAccessToken().tenant_access_token;
    let url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${bitable_app_token}/tables/${bitable_table_id}/records`;
    let options = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
            'Authorization': 'Bearer ' + tenant_access_token
        },
        'payload': JSON.stringify({ fields: payload })
    };

    try {
        let response = UrlFetchApp.fetch(url, options);
        let data = JSON.parse(response.getContentText());
        return data;
    } catch (error) {
        Logger.log('ERROR: Error in creating record: ' + error.toString());
        throw new Error('ERROR: Failed to create record.'); // 抛出异常确保不会标记为已读
    }
}
