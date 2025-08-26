-- conventions (a single event/brand; supports one upcoming edition)
CREATE TABLE conventions (
  id BIGINT UNSIGNED PRIMARY KEY AUTO_INCREMENT,
  name VARCHAR(255) NOT NULL,
  website_url VARCHAR(1024) NOT NULL,
  description TEXT NULL,
  venue VARCHAR(255) NULL,
  city VARCHAR(128) NULL,
  state VARCHAR(64) NULL,
  country VARCHAR(64) DEFAULT 'USA',
  timezone VARCHAR(64) DEFAULT 'America/Los_Angeles',
  start_date DATE NULL,
  end_date DATE NULL,
  next_edition_announced TINYINT(1) DEFAULT 0,
  logo_url VARCHAR(1024) NULL,
  source_url VARCHAR(1024) NULL,                -- where we last pulled core info
  last_scraped_at DATETIME NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  UNIQUE KEY uniq_con_name_site (name, website_url),
  KEY idx_dates (start_date, end_date),
  KEY idx_city_state (city, state)
);

-- normalized categories (comics, pop culture, anime, horror, movies, Disney, franchise-themed, etc.)
CREATE TABLE categories (
  id SMALLINT UNSIGNED PRIMARY KEY AUTO_INCREMENT,
  slug VARCHAR(64) UNIQUE NOT NULL,
  display_name VARCHAR(128) NOT NULL
);

-- many-to-many: conventions ↔ categories
CREATE TABLE convention_categories (
  convention_id BIGINT UNSIGNED NOT NULL,
  category_id SMALLINT UNSIGNED NOT NULL,
  PRIMARY KEY (convention_id, category_id),
  FOREIGN KEY (convention_id) REFERENCES conventions(id) ON DELETE CASCADE,
  FOREIGN KEY (category_id) REFERENCES categories(id) ON DELETE CASCADE
);

-- reference types of sign-up windows (attendee, press, pro, artist_alley, exhibitor, vendor, volunteer, panelist, etc.)
CREATE TABLE signup_types (
  id SMALLINT UNSIGNED PRIMARY KEY AUTO_INCREMENT,
  slug VARCHAR(64) UNIQUE NOT NULL,
  display_name VARCHAR(128) NOT NULL
);

-- per-convention sign-up links + windows
CREATE TABLE convention_signups (
  id BIGINT UNSIGNED PRIMARY KEY AUTO_INCREMENT,
  convention_id BIGINT UNSIGNED NOT NULL,
  signup_type_id SMALLINT UNSIGNED NOT NULL,
  link_url VARCHAR(1024) NOT NULL,
  open_at DATETIME NULL,
  close_at DATETIME NULL,
  status ENUM('unknown','announced','open','closed','waitlist') DEFAULT 'unknown',
  notes VARCHAR(1024) NULL,
  source_url VARCHAR(1024) NULL,                -- exact page we parsed
  last_scraped_at DATETIME NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  FOREIGN KEY (convention_id) REFERENCES conventions(id) ON DELETE CASCADE,
  FOREIGN KEY (signup_type_id) REFERENCES signup_types(id) ON DELETE RESTRICT,
  KEY idx_con_type (convention_id, signup_type_id),
  KEY idx_open_close (open_at, close_at)
);

-- auth (simple now; can expand to OAuth later)
CREATE TABLE users (
  id BIGINT UNSIGNED PRIMARY KEY AUTO_INCREMENT,
  email VARCHAR(320) UNIQUE NOT NULL,
  password_hash VARCHAR(255) NOT NULL,
  email_verified TINYINT(1) DEFAULT 0,
  default_lead_days INT DEFAULT 7,               -- user preference: remind N days before
  timezone VARCHAR(64) DEFAULT 'America/Los_Angeles',
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

-- user follows & reminder preferences per con/signup type
CREATE TABLE user_watchlist (
  id BIGINT UNSIGNED PRIMARY KEY AUTO_INCREMENT,
  user_id BIGINT UNSIGNED NOT NULL,
  convention_id BIGINT UNSIGNED NOT NULL,
  signup_type_id SMALLINT UNSIGNED NULL,         -- NULL = whole con (dates), else specific window
  lead_days INT NULL,                            -- override user default; e.g., 30 days before
  remind_on_open TINYINT(1) DEFAULT 1,           -- email at open time
  remind_before_open TINYINT(1) DEFAULT 1,       -- email N days before open
  remind_on_day_start TINYINT(1) DEFAULT 0,      -- for event start (if signup_type_id IS NULL)
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uniq_watch (user_id, convention_id, signup_type_id),
  FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
  FOREIGN KEY (convention_id) REFERENCES conventions(id) ON DELETE CASCADE,
  FOREIGN KEY (signup_type_id) REFERENCES signup_types(id) ON DELETE SET NULL
);

-- scheduled reminder instances (materialized so the cron can send)
CREATE TABLE reminders (
  id BIGINT UNSIGNED PRIMARY KEY AUTO_INCREMENT,
  user_id BIGINT UNSIGNED NOT NULL,
  convention_id BIGINT UNSIGNED NOT NULL,
  signup_type_id SMALLINT UNSIGNED NULL,
  trigger_at DATETIME NOT NULL,                  -- exact time to send (in UTC)
  kind ENUM('before_open','on_open','event_start') NOT NULL,
  sent_at DATETIME NULL,
  status ENUM('pending','sent','cancelled') DEFAULT 'pending',
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
  FOREIGN KEY (convention_id) REFERENCES conventions(id) ON DELETE CASCADE,
  FOREIGN KEY (signup_type_id) REFERENCES signup_types(id) ON DELETE SET NULL,
  KEY idx_trigger (status, trigger_at)
);

-- scraping sources + jobs (minimal, expandable)
CREATE TABLE sources (
  id BIGINT UNSIGNED PRIMARY KEY AUTO_INCREMENT,
  convention_id BIGINT UNSIGNED NULL,            -- may be NULL until matched
  kind ENUM('main','tickets','press','pro','artist','exhibitor','vendor','volunteer','schedule','faq','news') NOT NULL,
  url VARCHAR(1024) NOT NULL,
  enabled TINYINT(1) DEFAULT 1,
  last_checked_at DATETIME NULL,
  last_changed_at DATETIME NULL,
  etag VARCHAR(255) NULL,
  checksum CHAR(64) NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

CREATE TABLE scrape_jobs (
  id BIGINT UNSIGNED PRIMARY KEY AUTO_INCREMENT,
  source_id BIGINT UNSIGNED NOT NULL,
  run_at DATETIME NOT NULL,
  status ENUM('queued','running','success','error') DEFAULT 'queued',
  findings JSON NULL,                            -- structured extractions
  error TEXT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (source_id) REFERENCES sources(id) ON DELETE CASCADE,
  KEY idx_status_run (status, run_at)
);
