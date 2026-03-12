const fs = require('node:fs');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { DatabaseSync } = require('node:sqlite');

const FIXED_FIELD_VALUE_PREFIX = '__FIXED__:';

function normalizeText(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}

function parseJsonArray(rawValue) {
  try {
    const parsed = JSON.parse(rawValue);
    return Array.isArray(parsed) ? parsed : [];
  } catch (_error) {
    return [];
  }
}

function groupBigAccountRows(rows) {
  const groupMap = new Map();

  rows.forEach((row) => {
    const merchantId = normalizeText(row.merchantId);
    const currency = normalizeText(row.currency);

    if (!merchantId) {
      return;
    }

    if (!groupMap.has(merchantId)) {
      groupMap.set(merchantId, {
        merchantId,
        currencies: [],
        isMultiCurrency: false
      });
    }

    if (currency) {
      groupMap.get(merchantId).currencies.push(currency);
    }
  });

  return Array.from(groupMap.values()).map((item) => {
    const currencies = Array.from(new Set(item.currencies.filter((value) => value !== '')));
    return {
      merchantId: item.merchantId,
      currencies,
      isMultiCurrency: currencies.length > 1
    };
  });
}

class AppDatabase {
  constructor(dbPath) {
    this.dbPath = dbPath;
    this.db = null;
  }

  init() {
    fs.mkdirSync(path.dirname(this.dbPath), { recursive: true });
    this.db = new DatabaseSync(this.dbPath);
    this.db.exec('PRAGMA foreign_keys = ON;');
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS templates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_key TEXT,
        name TEXT NOT NULL UNIQUE,
        source_file_name TEXT NOT NULL,
        headers_json TEXT NOT NULL,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS template_mappings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_id INTEGER NOT NULL,
        template_field TEXT NOT NULL,
        mapped_field TEXT NOT NULL,
        row_index INTEGER NOT NULL,
        created_at TEXT NOT NULL,
        FOREIGN KEY (template_id) REFERENCES templates(id) ON DELETE CASCADE,
        UNIQUE(template_id, row_index)
      );

      CREATE TABLE IF NOT EXISTS template_big_accounts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_id INTEGER NOT NULL,
        merchant_id TEXT NOT NULL,
        currency TEXT NOT NULL,
        row_index INTEGER NOT NULL,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        FOREIGN KEY (template_id) REFERENCES templates(id) ON DELETE CASCADE,
        UNIQUE(template_id, merchant_id, currency)
      );

      CREATE TABLE IF NOT EXISTS app_settings (
        setting_key TEXT PRIMARY KEY,
        setting_value TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS account_mappings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        bank_account_id TEXT NOT NULL UNIQUE,
        clearing_account_id TEXT NOT NULL,
        row_index INTEGER NOT NULL UNIQUE,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );
    `);

    this.ensureTemplateKeySupport();
  }

  hasColumn(tableName, columnName) {
    const columns = this.db.prepare(`PRAGMA table_info(${tableName})`).all();
    return columns.some((column) => normalizeText(column.name) === normalizeText(columnName));
  }

  ensureTemplateKeySupport() {
    this.db.exec('BEGIN');

    try {
      if (!this.hasColumn('templates', 'template_key')) {
        this.db.exec('ALTER TABLE templates ADD COLUMN template_key TEXT;');
      }

      const rows = this.db
        .prepare(`
          SELECT id
          FROM templates
          WHERE COALESCE(template_key, '') = ''
        `)
        .all();

      const updateStatement = this.db.prepare(`
        UPDATE templates
        SET template_key = ?
        WHERE id = ?
      `);

      rows.forEach((row) => {
        updateStatement.run(randomUUID(), row.id);
      });

      this.db.exec(`
        CREATE UNIQUE INDEX IF NOT EXISTS templates_template_key_unique
        ON templates(template_key);
      `);

      this.db.exec('COMMIT');
    } catch (error) {
      this.db.exec('ROLLBACK');
      throw error;
    }
  }

  buildTemplateSummaryFromRow(row) {
    const merchantIdMappedField = normalizeText(row.merchantIdMappedField);
    const bigAccountCount = Number(row.bigAccountCount || 0);
    let bigAccountSummary = '未设置';
    let bigAccountMode = 'unset';

    if (merchantIdMappedField.startsWith(FIXED_FIELD_VALUE_PREFIX)) {
      if (bigAccountCount > 0) {
        bigAccountSummary = `${bigAccountCount}个`;
        bigAccountMode = 'multiple';
      } else {
        bigAccountSummary = '1个';
        bigAccountMode = 'single';
      }
    } else if (merchantIdMappedField) {
      bigAccountSummary = '来自账单';
      bigAccountMode = 'from-source';
    }

    return {
      id: row.id,
      templateKey: normalizeText(row.templateKey),
      name: row.name,
      sourceFileName: row.sourceFileName,
      headers: parseJsonArray(row.headersJson),
      createdAt: row.createdAt,
      updatedAt: row.updatedAt,
      mappingCount: Number(row.mappingCount || 0),
      bigAccountCount,
      bigAccountMode,
      bigAccountSummary
    };
  }

  listTemplates() {
    const statement = this.db.prepare(`
      SELECT
        t.id,
        t.template_key AS templateKey,
        t.name,
        t.source_file_name AS sourceFileName,
        t.headers_json AS headersJson,
        t.created_at AS createdAt,
        t.updated_at AS updatedAt,
        COUNT(DISTINCT m.id) AS mappingCount,
        COUNT(DISTINCT ba.merchant_id) AS bigAccountCount,
        merchant_mapping.mapped_field AS merchantIdMappedField
      FROM templates t
      LEFT JOIN template_mappings m ON m.template_id = t.id
      LEFT JOIN template_big_accounts ba ON ba.template_id = t.id
      LEFT JOIN template_mappings merchant_mapping
        ON merchant_mapping.template_id = t.id
        AND merchant_mapping.template_field = 'MerchantId'
      GROUP BY t.id
      ORDER BY t.updated_at DESC, t.id DESC
    `);

    return statement.all().map((row) => this.buildTemplateSummaryFromRow(row));
  }

  getTemplate(templateId) {
    const row = this.db
      .prepare(`
        SELECT
          t.id,
          t.template_key AS templateKey,
          t.name,
          t.source_file_name AS sourceFileName,
          t.headers_json AS headersJson,
          t.created_at AS createdAt,
          t.updated_at AS updatedAt,
          (
            SELECT COUNT(1)
            FROM template_mappings m
            WHERE m.template_id = t.id
          ) AS mappingCount,
          (
            SELECT COUNT(DISTINCT merchant_id)
            FROM template_big_accounts ba
            WHERE ba.template_id = t.id
          ) AS bigAccountCount,
          (
            SELECT mapped_field
            FROM template_mappings merchant_mapping
            WHERE merchant_mapping.template_id = t.id
              AND merchant_mapping.template_field = 'MerchantId'
            LIMIT 1
          ) AS merchantIdMappedField
        FROM templates t
        WHERE t.id = ?
      `)
      .get(templateId);

    return row ? this.buildTemplateSummaryFromRow(row) : null;
  }

  getTemplateByKey(templateKey) {
    const row = this.db
      .prepare(`
        SELECT id
        FROM templates
        WHERE template_key = ?
      `)
      .get(templateKey);

    return row ? this.getTemplate(row.id) : null;
  }

  getTemplateByName(name) {
    const row = this.db
      .prepare(`
        SELECT id
        FROM templates
        WHERE name = ?
      `)
      .get(name);

    return row ? this.getTemplate(row.id) : null;
  }

  upsertTemplate({ templateKey = '', name, sourceFileName, headers }) {
    const normalizedTemplateKey = normalizeText(templateKey) || randomUUID();
    const now = new Date().toISOString();
    const existingByKey = normalizeText(templateKey)
      ? this.db.prepare('SELECT id FROM templates WHERE template_key = ?').get(normalizedTemplateKey)
      : null;
    const existingByName = this.db.prepare('SELECT id FROM templates WHERE name = ?').get(name);
    const existing = existingByKey || existingByName;

    if (existing) {
      this.db.exec('BEGIN');
      try {
        this.db
          .prepare(`
            UPDATE templates
            SET template_key = ?, name = ?, source_file_name = ?, headers_json = ?, updated_at = ?
            WHERE id = ?
          `)
          .run(normalizedTemplateKey, name, sourceFileName, JSON.stringify(headers), now, existing.id);
        this.db.prepare('DELETE FROM template_mappings WHERE template_id = ?').run(existing.id);
        this.db.prepare('DELETE FROM template_big_accounts WHERE template_id = ?').run(existing.id);
        this.db.exec('COMMIT');
      } catch (error) {
        this.db.exec('ROLLBACK');
        throw error;
      }

      return this.getTemplate(existing.id);
    }

    const result = this.db
      .prepare(`
        INSERT INTO templates (template_key, name, source_file_name, headers_json, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?)
      `)
      .run(normalizedTemplateKey, name, sourceFileName, JSON.stringify(headers), now, now);

    return this.getTemplate(result.lastInsertRowid);
  }

  renameTemplate(templateId, nextName) {
    const now = new Date().toISOString();
    this.db
      .prepare(`
        UPDATE templates
        SET name = ?, updated_at = ?
        WHERE id = ?
      `)
      .run(nextName, now, templateId);

    return this.getTemplate(templateId);
  }

  deleteTemplate(templateId) {
    this.db.prepare('DELETE FROM templates WHERE id = ?').run(templateId);
  }

  getTemplateBigAccounts(templateId) {
    return this.db
      .prepare(`
        SELECT
          merchant_id AS merchantId,
          currency,
          row_index AS rowIndex
        FROM template_big_accounts
        WHERE template_id = ?
        ORDER BY row_index ASC, id ASC
      `)
      .all(templateId)
      .map((row) => ({
        merchantId: normalizeText(row.merchantId),
        currency: normalizeText(row.currency),
        rowIndex: Number(row.rowIndex || 0)
      }));
  }

  getTemplateMappings(templateId) {
    const template = this.getTemplate(templateId);

    if (!template) {
      return null;
    }

    const mappings = this.db
      .prepare(`
        SELECT row_index AS rowIndex, template_field AS templateField, mapped_field AS mappedField
        FROM template_mappings
        WHERE template_id = ?
        ORDER BY row_index ASC
      `)
      .all(templateId);
    const bigAccountRows = this.getTemplateBigAccounts(templateId);

    return {
      template,
      mappings,
      bigAccounts: groupBigAccountRows(bigAccountRows)
    };
  }

  saveMappings(templateId, mappings, bigAccounts = []) {
    const now = new Date().toISOString();
    this.db.exec('BEGIN');

    try {
      this.db.prepare('DELETE FROM template_mappings WHERE template_id = ?').run(templateId);
      this.db.prepare('DELETE FROM template_big_accounts WHERE template_id = ?').run(templateId);

      const insertMappingStatement = this.db.prepare(`
        INSERT INTO template_mappings (
          template_id, template_field, mapped_field, row_index, created_at
        ) VALUES (?, ?, ?, ?, ?)
      `);
      const insertBigAccountStatement = this.db.prepare(`
        INSERT INTO template_big_accounts (
          template_id, merchant_id, currency, row_index, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?)
      `);

      mappings.forEach((mapping, index) => {
        insertMappingStatement.run(
          templateId,
          mapping.templateField,
          mapping.mappedField,
          index,
          now
        );
      });

      bigAccounts.forEach((item, index) => {
        insertBigAccountStatement.run(
          templateId,
          item.merchantId,
          item.currency,
          index,
          now,
          now
        );
      });

      this.db
        .prepare('UPDATE templates SET updated_at = ? WHERE id = ?')
        .run(now, templateId);

      this.db.exec('COMMIT');
    } catch (error) {
      this.db.exec('ROLLBACK');
      throw error;
    }
  }

  listTemplateBundleEntries() {
    return this.listTemplates().map((template) => {
      const payload = this.getTemplateMappings(template.id);
      return {
        templateKey: template.templateKey,
        name: template.name,
        sourceFileName: template.sourceFileName,
        headers: template.headers,
        mappings: payload ? payload.mappings.map((mapping) => ({ ...mapping })) : [],
        bigAccounts: payload ? payload.bigAccounts.map((item) => ({
          merchantId: item.merchantId,
          currencies: item.currencies.slice(),
          isMultiCurrency: Boolean(item.isMultiCurrency)
        })) : [],
        createdAt: template.createdAt,
        updatedAt: template.updatedAt
      };
    });
  }

  getSetting(settingKey) {
    const row = this.db
      .prepare(`
        SELECT setting_value AS settingValue
        FROM app_settings
        WHERE setting_key = ?
      `)
      .get(settingKey);

    return row ? row.settingValue : null;
  }

  setSetting(settingKey, settingValue) {
    const now = new Date().toISOString();
    this.db
      .prepare(`
        INSERT INTO app_settings (setting_key, setting_value, updated_at)
        VALUES (?, ?, ?)
        ON CONFLICT(setting_key) DO UPDATE
        SET setting_value = excluded.setting_value,
            updated_at = excluded.updated_at
      `)
      .run(settingKey, settingValue, now);
  }

  getEnumConfig() {
    const raw = this.getSetting('enum_config');

    if (!raw) {
      return null;
    }

    try {
      return JSON.parse(raw);
    } catch (_error) {
      return null;
    }
  }

  setEnumConfig(enumConfig) {
    this.setSetting('enum_config', JSON.stringify(enumConfig));
  }

  getBackgroundConfig() {
    const raw = this.getSetting('background_config');

    if (!raw) {
      return null;
    }

    try {
      return JSON.parse(raw);
    } catch (_error) {
      return null;
    }
  }

  setBackgroundConfig(backgroundConfig) {
    this.setSetting('background_config', JSON.stringify(backgroundConfig));
  }

  listAccountMappings() {
    return this.db
      .prepare(`
        SELECT
          id,
          bank_account_id AS bankAccountId,
          clearing_account_id AS clearingAccountId,
          row_index AS rowIndex
        FROM account_mappings
        ORDER BY row_index ASC, id ASC
      `)
      .all();
  }

  saveAccountMappings(mappings) {
    const now = new Date().toISOString();
    this.db.exec('BEGIN');

    try {
      this.db.exec('DELETE FROM account_mappings');

      const insertStatement = this.db.prepare(`
        INSERT INTO account_mappings (
          bank_account_id, clearing_account_id, row_index, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?)
      `);

      mappings.forEach((mapping, index) => {
        insertStatement.run(
          mapping.bankAccountId,
          mapping.clearingAccountId,
          index,
          now,
          now
        );
      });

      this.db.exec('COMMIT');
    } catch (error) {
      this.db.exec('ROLLBACK');
      throw error;
    }
  }
}

module.exports = {
  AppDatabase
};
