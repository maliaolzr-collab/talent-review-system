-- ============================================================
-- Supabase Migration Script
-- 在 Supabase SQL Editor 中执行此脚本
-- ============================================================

-- 1. 部门表
CREATE TABLE IF NOT EXISTS departments (
    id SERIAL PRIMARY KEY,
    category TEXT NOT NULL,
    sub_dept TEXT NOT NULL,
    UNIQUE(category, sub_dept)
);

-- 2. 员工表
CREATE TABLE IF NOT EXISTS employees (
    id SERIAL PRIMARY KEY,
    dept_id INTEGER NOT NULL REFERENCES departments(id),
    chinese_name TEXT,
    english_name TEXT,
    position_title TEXT,
    job_responsibility TEXT,
    job_level TEXT,
    age INTEGER,
    education TEXT,
    graduation_institution TEXT,
    graduation_date TEXT,
    work_experience TEXT,
    entry_date TEXT,
    company_tenure TEXT,
    base_salary DOUBLE PRECISION,
    performance_salary DOUBLE PRECISION,
    total_salary DOUBLE PRECISION,
    knowledge_skill_match DOUBLE PRECISION,
    problem_solving_match DOUBLE PRECISION,
    responsibility_match DOUBLE PRECISION,
    person_position_score DOUBLE PRECISION,
    annual_performance TEXT,
    learning_ability INTEGER,
    thinking_ability INTEGER,
    understanding_others INTEGER,
    emotional_maturity INTEGER,
    potential_score DOUBLE PRECISION,
    performance_level TEXT,
    potential_level TEXT,
    grid_position INTEGER,
    grid_name TEXT,
    talent_pipeline TEXT,
    result_application TEXT,
    development_plan TEXT,
    management_strategy TEXT,
    upload_batch_id INTEGER,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 3. 用户表
CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL,
    dept_category TEXT,
    display_name TEXT
);

-- 4. 上传批次表
CREATE TABLE IF NOT EXISTS upload_batches (
    id SERIAL PRIMARY KEY,
    dept_category TEXT NOT NULL,
    sub_dept TEXT NOT NULL,
    filename TEXT,
    uploaded_by INTEGER,
    employee_count INTEGER,
    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 索引
CREATE INDEX IF NOT EXISTS idx_employees_dept ON employees(dept_id);
CREATE INDEX IF NOT EXISTS idx_employees_grid ON employees(grid_position);
CREATE INDEX IF NOT EXISTS idx_employees_pipeline ON employees(talent_pipeline);
CREATE INDEX IF NOT EXISTS idx_upload_batches_dept ON upload_batches(dept_category);
