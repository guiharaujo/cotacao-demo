-- =============================================================================
-- QUOTATION SYSTEM — Azure SQL Schema (Demo)
-- =============================================================================
-- All table names use the COT_ prefix.
-- Deploy this on Azure SQL before connecting the workbook.
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Users
-- -----------------------------------------------------------------------------
CREATE TABLE COT_USUARIOS (
    ID          INT IDENTITY(1,1) PRIMARY KEY,
    USUARIO     NVARCHAR(100) NOT NULL UNIQUE,   -- login
    SENHA       NVARCHAR(255) NOT NULL,           -- hashed in production
    USERNAME    NVARCHAR(150) NOT NULL,           -- display name
    CARGO       NVARCHAR(100),                    -- role (e.g. Seller, Management)
    DEPARTAMENTO NVARCHAR(100),
    EMAIL       NVARCHAR(200),
    CELULAR     NVARCHAR(20),
    ATIVO       NVARCHAR(3) NOT NULL DEFAULT 'SIM'  -- SIM / NAO
);

-- Sample users (replace passwords with hashed values in production)
INSERT INTO COT_USUARIOS (USUARIO, SENHA, USERNAME, CARGO, DEPARTAMENTO, EMAIL, CELULAR, ATIVO)
VALUES
    ('john.doe',    'demo1234', 'John Doe',    'Seller',     'Commercial', 'john@example.com',  '11999990001', 'SIM'),
    ('jane.smith',  'demo1234', 'Jane Smith',  'Management', 'Commercial', 'jane@example.com',  '11999990002', 'SIM'),
    ('admin',       'admin123', 'Admin User',  'Admin',      'IT',         'admin@example.com', '11999990003', 'SIM');

-- -----------------------------------------------------------------------------
-- Products
-- -----------------------------------------------------------------------------
CREATE TABLE COT_PRODUTOS (
    ID                  INT IDENTITY(1,1) PRIMARY KEY,
    STATUS              NVARCHAR(10)  NOT NULL DEFAULT 'ACTIVE',  -- ACTIVE / INACTIVE
    CODE                NVARCHAR(50)  NOT NULL UNIQUE,
    DESCRIPTION         NVARCHAR(255) NOT NULL,
    NCM                 NVARCHAR(20),
    FAMILY              NVARCHAR(100),
    WATTAGE             DECIMAL(10,2),
    COLOR_TEMP          NVARCHAR(20),
    TYPE                NVARCHAR(50),
    COST                DECIMAL(18,4) NOT NULL,   -- unit cost (R$)
    IE                  NVARCHAR(5),              -- IE exemption flag
    WEIGHT              DECIMAL(10,3),            -- kg
    HEIGHT              DECIMAL(10,2),            -- cm
    WIDTH               DECIMAL(10,2),
    LENGTH              DECIMAL(10,2),
    SHARED_BOX          INT,                      -- units per shared box
    SEG                 NVARCHAR(20),             -- market segment
    ICMS                DECIMAL(5,4),             -- ICMS rate (e.g. 0.12)
    PIS_COFINS          DECIMAL(5,4),
    IPI                 DECIMAL(5,4),
    COMMISSION          DECIMAL(5,4),
    LOGISTICS           DECIMAL(5,4),
    OPERATIONAL         DECIMAL(5,4),
    SALES               DECIMAL(5,4),
    ADMIN               DECIMAL(5,4),
    QUALITY             DECIMAL(5,4),
    BENEFIT             DECIMAL(5,4),
    DIFAL               DECIMAL(5,4),
    CONTRIBUTION_MARGIN DECIMAL(5,4)              -- minimum target margin
);

-- Sample products (fictional data)
INSERT INTO COT_PRODUTOS
    (STATUS, CODE, DESCRIPTION, NCM, FAMILY, WATTAGE, COLOR_TEMP, TYPE,
     COST, WEIGHT, HEIGHT, WIDTH, LENGTH, SHARED_BOX,
     SEG, ICMS, PIS_COFINS, IPI, COMMISSION, LOGISTICS,
     OPERATIONAL, SALES, ADMIN, QUALITY, BENEFIT, DIFAL, CONTRIBUTION_MARGIN)
VALUES
    ('ACTIVE','LED-100W-4K','LED Fixture 100W 4000K','85395000','Industrial',100,'4000K','Fixture',
     320.00, 4.5, 12, 30, 60, 4, 'Public', 0.12, 0.0925, 0.05, 0.05, 0.02, 0.02, 0.03, 0.02, 0.01, 0.00, 0.0200, 0.30),

    ('ACTIVE','LED-150W-5K','LED Fixture 150W 5000K','85395000','Industrial',150,'5000K','Fixture',
     480.00, 6.2, 14, 35, 70, 4, 'Public', 0.12, 0.0925, 0.05, 0.05, 0.02, 0.02, 0.03, 0.02, 0.01, 0.00, 0.0200, 0.30),

    ('ACTIVE','LED-50W-3K','LED Panel 50W 3000K','85395000','Commercial', 50,'3000K','Panel',
     180.00, 2.1,  4, 60, 60, 6, 'Private', 0.12, 0.0925, 0.00, 0.05, 0.02, 0.02, 0.03, 0.02, 0.01, 0.00, 0.0200, 0.28),

    ('ACTIVE','LED-200W-4K','LED HighBay 200W 4000K','85395000','Industrial',200,'4000K','HighBay',
     620.00, 8.0, 18, 40, 40, 2, 'Public', 0.12, 0.0925, 0.05, 0.05, 0.02, 0.02, 0.03, 0.02, 0.01, 0.00, 0.0200, 0.32);

-- -----------------------------------------------------------------------------
-- Payment terms
-- -----------------------------------------------------------------------------
CREATE TABLE COT_PAGAMENTO (
    ID          INT IDENTITY(1,1) PRIMARY KEY,
    DESCRIPTION NVARCHAR(100) NOT NULL,
    RATE        DECIMAL(5,4)  NOT NULL DEFAULT 0  -- financial cost rate
);

INSERT INTO COT_PAGAMENTO (DESCRIPTION, RATE) VALUES
    ('Cash on delivery',     0.0000),
    ('30 days',              0.0120),
    ('30/60 days',           0.0240),
    ('30/60/90 days',        0.0360),
    ('30/60/90/120 days',    0.0480);

-- -----------------------------------------------------------------------------
-- Proposal audit log
-- -----------------------------------------------------------------------------
CREATE TABLE COT_PROPOSTAS (
    ID              INT IDENTITY(1,1) PRIMARY KEY,
    NUMERO          NVARCHAR(20)   NOT NULL,   -- sequential proposal number
    DATA_GERACAO    DATETIME       NOT NULL DEFAULT GETDATE(),
    USUARIO         NVARCHAR(150),
    CARGO           NVARCHAR(100),
    RAZAO_SOCIAL    NVARCHAR(255),
    CNPJ            NVARCHAR(20),
    TOTAL           DECIMAL(18,2),
    MARGEM_PCT      DECIMAL(5,2),
    ALCADA          NVARCHAR(50),
    OPP_CRM         NVARCHAR(50)
);

-- -----------------------------------------------------------------------------
-- Sequential proposal counter
-- -----------------------------------------------------------------------------
CREATE TABLE COT_CONTADOR (
    ID       INT IDENTITY(1,1) PRIMARY KEY,
    CHAVE    NVARCHAR(50) NOT NULL UNIQUE,
    VALOR    INT NOT NULL DEFAULT 0
);

INSERT INTO COT_CONTADOR (CHAVE, VALOR) VALUES ('PROPOSTA_SEQ', 1000);
