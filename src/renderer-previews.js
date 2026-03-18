(function initRendererPreviews(global) {
  function createRendererPreviews(deps) {
    const {
      state,
      elements,
      MODULES,
      ADVANCED_MAPPING_FIELDS,
      BALANCE_CALCULATED_OPTION,
      MERCHANT_ID_SELF_INPUT_OPTION,
      SIGNED_AMOUNT_MAPPING_FIELD,
      AMOUNT_BASED_NAME_MAPPING_FIELD,
      AMOUNT_BASED_ACCOUNT_MAPPING_FIELD,
      setCurrentModule,
      syncNewAccountCurrencyMode,
      updateNewAccountGenerateAvailability,
      setNewAccountExportAvailability,
      setNewAccountStatus,
      setExportAvailability,
      setStatus,
      getNewAccountStatusTitle,
      setNewAccountOpenDateValue,
      openModal,
      createTemplateManagerDialog,
      createMappingDialog,
      createTemplateRenameDialog,
      createBigAccountManagerDialog,
      createBigAccountSelectionDialog,
      closeModal,
      openBackgroundPalette
    } = deps;

    function applyNewAccountPreviewState() {
      setCurrentModule(MODULES.newAccountGenerator.id);
      elements.newAccountMultiCurrencyCheckbox.checked = false;
      state.selectedNewAccountCurrencies = [];
      syncNewAccountCurrencyMode();
      elements.newAccountBankNameInput.value = '中国银行';
      elements.newAccountLocationInput.value = '香港';
      elements.newAccountCurrencyInput.value = 'USD';
      elements.newAccountBankAccountInput.value = '6222000000000001';
      setNewAccountOpenDateValue('2026-01-01');
      updateNewAccountGenerateAvailability();
      setNewAccountExportAvailability(true);
      setNewAccountStatus('新开账户余额账单可导出', 'success', {
        errorReportReady: false,
        idleTitle: getNewAccountStatusTitle()
      });
    }

    function applyTemplateManagerPreviewState() {
      setCurrentModule(MODULES.statementGenerator.id);
      state.templates = [
        {
          id: 'preview-template-1',
          name: 'LusoBank-MO',
          bigAccountSummary: '来自账单'
        },
        {
          id: 'preview-template-2',
          name: 'BankABC-HK',
          bigAccountSummary: '未设置'
        },
        {
          id: 'preview-template-3',
          name: 'PingPong-US',
          bigAccountSummary: '62220000000000012345'
        },
        {
          id: 'preview-template-4',
          name: 'HSBC-SG',
          bigAccountSummary: '3个'
        }
      ];
      openModal(createTemplateManagerDialog());
    }

    function buildPreviewMappingPayload() {
      return {
        template: {
          id: 'preview-template-4',
          name: 'HSBC-SG',
          headers: [
            '交易日期',
            '起息日期',
            '发生额',
            '余额',
            '对手户名',
            '对手账号',
            '币种',
            '附言'
          ]
        },
        targetFields: [
          'BillDate',
          'ValueDate',
          'Credit Amount',
          'Debit Amount',
          'Balance',
          'MerchantId',
          'Currency',
          'Payee Name',
          'Payee Cardno',
          'Drawee Name',
          'Drawee CardNo',
          SIGNED_AMOUNT_MAPPING_FIELD,
          AMOUNT_BASED_NAME_MAPPING_FIELD,
          AMOUNT_BASED_ACCOUNT_MAPPING_FIELD
        ],
        mappings: [
          { templateField: 'BillDate', mappedField: '交易日期', customValue: '', isMultiBigAccount: false },
          { templateField: 'ValueDate', mappedField: '起息日期', customValue: '', isMultiBigAccount: false },
          { templateField: 'Credit Amount', mappedField: '', customValue: '', isMultiBigAccount: false },
          { templateField: 'Debit Amount', mappedField: '', customValue: '', isMultiBigAccount: false },
          { templateField: 'Balance', mappedField: BALANCE_CALCULATED_OPTION, customValue: '', isMultiBigAccount: false },
          { templateField: 'MerchantId', mappedField: MERCHANT_ID_SELF_INPUT_OPTION, customValue: '', isMultiBigAccount: true },
          { templateField: 'Currency', mappedField: '', customValue: '', isMultiBigAccount: false },
          { templateField: 'Payee Name', mappedField: '', customValue: '', isMultiBigAccount: false },
          { templateField: 'Payee Cardno', mappedField: '', customValue: '', isMultiBigAccount: false },
          { templateField: 'Drawee Name', mappedField: '', customValue: '', isMultiBigAccount: false },
          { templateField: 'Drawee CardNo', mappedField: '', customValue: '', isMultiBigAccount: false },
          { templateField: SIGNED_AMOUNT_MAPPING_FIELD, mappedField: '发生额', customValue: '', isMultiBigAccount: false },
          { templateField: AMOUNT_BASED_NAME_MAPPING_FIELD, mappedField: '对手户名', customValue: '', isMultiBigAccount: false },
          { templateField: AMOUNT_BASED_ACCOUNT_MAPPING_FIELD, mappedField: '对手账号', customValue: '', isMultiBigAccount: false }
        ],
        bigAccounts: [
          {
            merchantId: '6222000000000001',
            currencies: ['USD'],
            isMultiBigAccount: false
          },
          {
            merchantId: '6222000000000001',
            currencies: ['HKD', 'CNY', 'EUR'],
            isMultiBigAccount: true
          }
        ],
        advancedMappingFields: ADVANCED_MAPPING_FIELDS.slice()
      };
    }

    function applyMappingDialogPreviewState() {
      setCurrentModule(MODULES.statementGenerator.id);
      state.currencyOptions = ['USD', 'HKD', 'CNY', 'EUR', 'JPY'];
      openModal(createMappingDialog(buildPreviewMappingPayload()));
    }

    function applyTemplateRenamePreviewState() {
      setCurrentModule(MODULES.statementGenerator.id);
      openModal(createTemplateRenameDialog({
        id: 'preview-template-2',
        name: 'BankABC-HK'
      }));
    }

    function applyBigAccountManagerPreviewState() {
      setCurrentModule(MODULES.statementGenerator.id);
      state.currencyOptions = ['USD', 'HKD', 'CNY', 'EUR', 'JPY'];
      openModal(createBigAccountManagerDialog({
        bigAccounts: [
          {
            merchantId: '6222000000000001',
            currencies: ['USD'],
            isMultiCurrency: false
          },
          {
            merchantId: '6222000000000001',
            currencies: ['HKD', 'CNY', 'EUR', 'JPY'],
            isMultiCurrency: true
          },
          {
            merchantId: '9558800000000008',
            currencies: ['SGD', 'USD'],
            isMultiCurrency: true
          }
        ],
        onDone: () => {},
        onCancel: closeModal
      }));

      setTimeout(() => {
        const addButton = elements.modalRoot.querySelector('.big-account-card [data-action="add"]');
        addButton?.click();
        const rows = Array.from(elements.modalRoot.querySelectorAll('tr[data-big-account-row]'));
        const lastRow = rows[rows.length - 1];
        if (!lastRow) {
          return;
        }

        const merchantInput = lastRow.querySelector('.big-account-merchant-input');
        const currencySelect = lastRow.querySelector('.big-account-currency-select');
        if (merchantInput) {
          merchantInput.value = '8888999900001111';
        }

        if (currencySelect) {
          currencySelect.value = 'USD';
          currencySelect.dispatchEvent(new Event('change'));
        }
      }, 40);
    }

    function applyBigAccountManagerDropdownPreviewState() {
      applyBigAccountManagerPreviewState();

      setTimeout(() => {
        const rows = Array.from(elements.modalRoot.querySelectorAll('tr[data-big-account-row]'));
        const targetRow = rows[1];

        if (!targetRow) {
          return;
        }

        targetRow.querySelector('[data-action="toggle-complete"]')?.click();
        targetRow.querySelector('.big-account-currency-dropdown-btn')?.click();
      }, 160);
    }

    function applyBigAccountSelectionPreviewState() {
      setCurrentModule(MODULES.statementGenerator.id);
      openModal(createBigAccountSelectionDialog([
        {
          label: '6222000000000001 / USD',
          merchantId: '6222000000000001',
          currency: 'USD'
        },
        {
          label: '6222000000000001 / HKD',
          merchantId: '6222000000000001',
          currency: 'HKD'
        },
        {
          label: '9558800000000008 / SGD',
          merchantId: '9558800000000008',
          currency: 'SGD'
        }
      ]));

      setTimeout(() => {
        const firstOption = elements.modalRoot.querySelector('.big-account-selection-list input[type="radio"]');
        if (firstOption) {
          firstOption.checked = true;
        }
      }, 40);
    }

    return {
      applyNewAccountPreviewState,
      applyTemplateManagerPreviewState,
      buildPreviewMappingPayload,
      applyMappingDialogPreviewState,
      applyTemplateRenamePreviewState,
      applyBigAccountManagerPreviewState,
      applyBigAccountManagerDropdownPreviewState,
      applyBigAccountSelectionPreviewState
    };
  }

  global.__rendererPreviews = {
    createRendererPreviews
  };
}(window));
