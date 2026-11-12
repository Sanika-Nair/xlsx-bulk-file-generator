VerifyPayments()
{
    /* How the app works:
       page 1 shows up to 1000 rows -> Select All + Verify + OK
       -> verified rows drop out, page 1 reloads with the next 1000 -> repeat
       until no rows are left. filter.pageNumber is always 1.
       e.g. 18,524 records = 18 batches of 1000 + 1 batch of 524.

       The first batch of rows + token were captured by Click_List.            */

    srand((unsigned int)time(NULL));
    g_totalVerified = 0;

    for (g_batch = 1; g_batch <= MAX_BATCHES; g_batch++) {

        /* ---------------- rows on the current page ---------------- */
        g_rows    = atoi(lr_eval_string("{PayInsCode_count}"));
        g_cntAmt  = atoi(lr_eval_string("{PayAmount_count}"));
        g_cntBene = atoi(lr_eval_string("{BeneName_count}"));

        if (g_rows == 0) {
            lr_output_message("All records verified: %d rows in %d batch(es)",
                              g_totalVerified, g_batch - 1);
            break;
        }

        if (g_cntAmt != g_rows || g_cntBene != g_rows) {
            lr_error_message("Batch %d: field counts differ (code=%d amt=%d bene=%d) - check boundaries",
                             g_batch, g_rows, g_cntAmt, g_cntBene);
            lr_exit(LR_EXIT_ITERATION_AND_CONTINUE, LR_FAIL);
        }

        if (g_rows > BATCH_SIZE) {
            lr_error_message("Batch %d: %d rows captured (more than %d) - boundaries match too much",
                             g_batch, g_rows, BATCH_SIZE);
            lr_exit(LR_EXIT_ITERATION_AND_CONTINUE, LR_FAIL);
        }

        /* ---------------- build paymentList for this batch ---------------- */
        g_cap  = 64;
        g_used = 0;

        for (g_i = 1; g_i <= g_rows; g_i++)
            g_cap += strlen(lr_paramarr_idx("PayInsCode", g_i))
                   + strlen(lr_paramarr_idx("PayAmount",  g_i))
                   + strlen(lr_paramarr_idx("BeneName",   g_i))
                   + 110;                              /* keys, quotes, commas, idx digits */

        g_buf = (char *)malloc(g_cap);
        if (g_buf == NULL) {
            lr_error_message("Batch %d: malloc failed (%d bytes)", g_batch, g_cap);
            lr_exit(LR_EXIT_ITERATION_AND_CONTINUE, LR_FAIL);
        }

        g_used += sprintf(g_buf + g_used, "{\"paymentList\":[");

        for (g_i = 1; g_i <= g_rows; g_i++) {
            g_used += sprintf(g_buf + g_used,
                "%s{\"paymentInsCode\":\"%s\",\"amount\":\"%s\",\"beneName\":\"%s\",\"reason\":\"\",\"idx\":\"%d\"}",
                (g_i > 1 ? "," : ""),
                lr_paramarr_idx("PayInsCode", g_i),
                lr_paramarr_idx("PayAmount",  g_i),
                lr_paramarr_idx("BeneName",   g_i),
                g_i - 1);                       /* idx = row position on page: 0..999 */
        }

        g_used += sprintf(g_buf + g_used, "]}");

        lr_save_string(g_buf, "PaymentListJson");
        free(g_buf);
        g_buf = NULL;

        /* ---------------- rand= for the URL ---------------- */
        sprintf(g_rand, "%.16f", (double)rand() / (double)RAND_MAX);
        lr_save_string(g_rand, "Rand");

        lr_output_message("Batch %d: verifying %d rows (remaining: %s, amount: %s)",
                          g_batch, g_rows,
                          lr_eval_string("{TotalTxn}"), lr_eval_string("{TotalAmt}"));

        /* ================= Select All + Verify ================= */

        /* token in the Verify response is needed by the OK/reload request */
        web_reg_save_param_ex(
            "ParamName=StrutsToken",
            "LB=name=\"struts.token\" value=\"",     /* TODO verify */
            "RB=\"",
            "NotFound=error",
            SEARCH_FILTERS, "Scope=Body", LAST);

        web_reg_find("Text=iframe", LAST);   /* TODO: better - the "records verified" popup text */

        lr_start_transaction("04_SelectAll_Verify");

        web_submit_data("verifyPaymentsLegAction.action",
            "Action=https://indiaeb.uat.dev.echonet/BNPPNetPayApp/verifyPaymentsLegAction.action?rand={Rand}",
            "Method=POST",
            "TargetFrame=",
            "RecContentType=text/html",
            "Referer=https://indiaeb.uat.dev.echonet/BNPPNetPayApp/verifyPaymentDetailsListAction.action",
            "Mode=HTML",
            ITEMDATA,
            "Name=screenType", "Value=VERIFY", ENDITEM,
            "Name=roleCode", "Value=", ENDITEM,
            "Name=editEnable", "Value=N", ENDITEM,
            "Name=filter.firstLoad", "Value=N", ENDITEM,
            "Name=filter.sortField", "Value=", ENDITEM,
            "Name=filter.orderType", "Value=", ENDITEM,
            "Name=filter.colNumber", "Value=", ENDITEM,
            "Name=filter.pageNumber", "Value=1", ENDITEM,
            "Name=filter.selectAll", "Value=N", ENDITEM,
            "Name=filter.sessionId", "Value=", ENDITEM,
            "Name=prefEmailListFrame", "Value=", ENDITEM,
            /* TODO: line 27 of your recording (hidden in the screenshot) goes here */
            "Name=sid", "Value={Sid}", ENDITEM,
            "Name=filter.fileName", "Value={FileName}", ENDITEM,
            "Name=filter.filterStatus", "Value=30", ENDITEM,
            "Name=filter.beneName", "Value=", ENDITEM,
            "Name=filter.accountNumber", "Value=", ENDITEM,
            "Name=filter.preferredAuthorizer", "Value=", ENDITEM,
            "Name=todayDate", "Value=", ENDITEM,
            "Name=filter.uploadDataFrom", "Value=", ENDITEM,
            "Name=filter.uploadDataTo", "Value=", ENDITEM,
            "Name=filter.amountFrom", "Value=", ENDITEM,
            "Name=filter.amountTo", "Value=", ENDITEM,
            "Name=filter.surveillanceStatus", "Value=", ENDITEM,
            "Name=filter.makerUserCode", "Value=-1", ENDITEM,
            "Name=filter.custRefNumber", "Value=", ENDITEM,
            "Name=filter.additionalInput", "Value=", ENDITEM,
            "Name=filter.perPage", "Value=1000", ENDITEM,
            "Name=filter.pageLoad", "Value=N", ENDITEM,
            "Name=VALUE_DATE", "Value={ValueDate}", ENDITEM,
            "Name=TD_CUTOFF_TIME", "Value=18:00", ENDITEM,
            "Name=CD_CUTOFF_TIME", "Value=18:00", ENDITEM,
            "Name=custRefNumber", "Value=", ENDITEM,
            "Name=additionalInput", "Value=", ENDITEM,
            "Name=additionalDetails", "Value=", ENDITEM,
            "Name=isFxInteractive", "Value=false", ENDITEM,
            "Name=total_txn", "Value={TotalTxn}", ENDITEM,
            "Name=total_amt", "Value={TotalAmt}", ENDITEM,
            "Name=shadowConfigBean", "Value={\"shadowConfigBean\":{\"isShadowPeriod\":false,\"shadowPeriodProducts\":[],\"isDirectSubToNetPay\":false}}", ENDITEM,
            "Name=struts.token.name", "Value=struts.token", ENDITEM,
            "Name=struts.token", "Value={StrutsToken}", ENDITEM,
            "Name=filter.actionType", "Value=verify_verify", ENDITEM,
            "Name=paymentList", "Value={PaymentListJson}", ENDITEM,
            "Name=rsacodetextbox", "Value=", ENDITEM,
            "Name=vascotextbox1", "Value=", ENDITEM,
            "Name=vascotextbox2", "Value=", ENDITEM,
            "Name=allAction", "Value=N", ENDITEM,
            LAST);

        lr_end_transaction("04_SelectAll_Verify", LR_AUTO);

        g_totalVerified += g_rows;

        /* ================= OK on popup -> page 1 reloads with next 1000 ================= */

        /* Captures for the reload response: next batch of rows + next token.
           Same boundaries as in Click_List - keep both copies in sync.        */
        web_reg_save_param_ex(
            "ParamName=StrutsToken",
            "LB=name=\"struts.token\" value=\"",
            "RB=\"",
            "NotFound=error",
            SEARCH_FILTERS, "Scope=Body", LAST);

        web_reg_save_param_ex(
            "ParamName=Sid",
            "LB=name=\"sid\" value=\"",
            "RB=\"",
            "NotFound=warning",
            SEARCH_FILTERS, "Scope=Body", LAST);

        web_reg_save_param_ex(
            "ParamName=TotalTxn",
            "LB=name=\"total_txn\" value=\"",
            "RB=\"",
            "NotFound=warning",
            SEARCH_FILTERS, "Scope=Body", LAST);

        web_reg_save_param_ex(
            "ParamName=TotalAmt",
            "LB=name=\"total_amt\" value=\"",
            "RB=\"",
            "NotFound=warning",
            SEARCH_FILTERS, "Scope=Body", LAST);

        web_reg_save_param_ex(
            "ParamName=PayInsCode",
            "LB=\"paymentInsCode\":\"",
            "RB=\"",
            "Ordinal=ALL", "NotFound=warning",
            SEARCH_FILTERS, "Scope=Body", LAST);

        web_reg_save_param_ex(
            "ParamName=PayAmount",
            "LB=\"amount\":\"",
            "RB=\"",
            "Ordinal=ALL", "NotFound=warning",
            SEARCH_FILTERS, "Scope=Body", LAST);

        web_reg_save_param_ex(
            "ParamName=BeneName",
            "LB=\"beneName\":\"",
            "RB=\"",
            "Ordinal=ALL", "NotFound=warning",
            SEARCH_FILTERS, "Scope=Body", LAST);

            web_reg_find("Text=iframe", LAST);

        lr_start_transaction("05_OK_LoadNextBatch");

        /* TODO: paste your recorded web_submit_data("paymentListingPaginateAction.action_3", ...)
                 here and replace in it:
                   struts.token       -> {StrutsToken}
                   sid                -> {Sid}
                   filter.fileName    -> {FileName}
                   filter.pageNumber  -> 1
                   total_txn          -> {TotalTxn}
                   total_amt          -> {TotalAmt}
                   VALUE_DATE         -> {ValueDate}
                   rand= (if in URL)  -> {Rand}
                 If its body also has a paymentList, use {PaymentListJson}.
                 Then delete all the other recorded verify / paginate pairs. */

        lr_end_transaction("05_OK_LoadNextBatch", LR_AUTO);
    }

    if (g_batch > MAX_BATCHES)
        lr_error_message("Stopped at MAX_BATCHES (%d) - rows still remaining", MAX_BATCHES);

    return 0;
}
