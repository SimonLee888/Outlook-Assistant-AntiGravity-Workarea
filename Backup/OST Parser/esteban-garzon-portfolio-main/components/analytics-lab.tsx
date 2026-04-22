"use client";

import { motion } from "framer-motion";
import {
  analyticsArchive,
  analyticsFeatured,
  analyticsSignals,
  profile,
} from "@/lib/portfolio-data";
import styles from "./code-lab.module.css";

export function AnalyticsLab() {
  return (
    <main className={styles.page}>
      <div className={styles.backdrop} />

      <header className={styles.header}>
        <a href="/" className={styles.brand}>
          {profile.name}
        </a>

        <div className={styles.headerLinks}>
          <a href="/">Home</a>
          <a href="/code">Code Lab</a>
          <a href="#archive">Archive</a>
          <a href={profile.linkedin} target="_blank" rel="noreferrer">
            LinkedIn
          </a>
        </div>
      </header>

      <section className={styles.hero}>
        <motion.div
          initial={{ opacity: 0, y: 24 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6 }}
        >
          <span className={styles.eyebrow}>Applied Analytics</span>
          <h1>Forecasting, business intelligence, graph analysis, visualization, and academic case work curated into one clean archive.</h1>
          <p>
            This page adds an analytics layer to the main portfolio without replacing
            the site&apos;s core identity as a systems, automation, and data-engineering
            portfolio. The strongest case studies stay visible first, while the broader
            archive keeps the range of work accessible and organized.
          </p>
        </motion.div>
      </section>

      <section className={styles.signalStrip}>
        {analyticsSignals.map((signal, index) => (
          <motion.article
            key={signal.label}
            className={styles.signalCard}
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true, amount: 0.4 }}
            transition={{ duration: 0.45, delay: index * 0.08 }}
          >
            <span>{signal.label}</span>
            <strong>{signal.value}</strong>
            <p>{signal.detail}</p>
          </motion.article>
        ))}
      </section>

      <section className={styles.trackGrid}>
        {analyticsFeatured.map((project, index) => (
          <motion.article
            key={project.title}
            className={styles.trackCard}
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true, amount: 0.3 }}
            transition={{ duration: 0.45, delay: index * 0.08 }}
          >
            <span className={styles.resourceTag}>{project.tag}</span>
            <h2>{project.title}</h2>
            <p>{project.summary}</p>

            <div className={styles.trackPoints}>
              <div className={styles.trackPoint}>{project.outcome}</div>
              {project.stack.map((item) => (
                <div key={item} className={styles.trackPoint}>
                  {item}
                </div>
              ))}
            </div>

            <div className={styles.resourceActions}>
              <a href={project.href} target={project.href.endsWith(".pdf") ? "_blank" : undefined} rel="noreferrer">
                {project.hrefLabel}
              </a>
            </div>
          </motion.article>
        ))}
      </section>

      <section className={styles.dualSection}>
        <article className={styles.panelCard}>
          <div className={styles.panelHeading}>
            <span className={styles.eyebrow}>Why this page exists</span>
            <h2>The home page stays selective, while this page carries the wider analytics footprint.</h2>
          </div>
          <div className={styles.requirementGrid}>
            <div className={styles.requirementCard}>
              <strong>No duplication of the main story</strong>
              <p>
                The core homepage remains centered on CTO-level systems, ETL, automation,
                and operational software rather than becoming a crowded coursework catalog.
              </p>
            </div>
            <div className={styles.requirementCard}>
              <strong>Everything stays in English</strong>
              <p>
                All analytics content is presented in English so it fits the public-facing
                portfolio and domain already deployed on Vercel.
              </p>
            </div>
          </div>
        </article>

        <article className={styles.panelCard}>
          <div className={styles.panelHeading}>
            <span className={styles.eyebrow}>What it adds</span>
            <h2>Forecasting, BI, graph analytics, visualization, databases, and optimization proof.</h2>
          </div>
          <div className={styles.workflowList}>
            <div className={styles.workflowItem}>
              <span>01</span>
              <div>
                <strong>Forecasting and planning</strong>
                <p>Thesis and strategy material show Esteban thinking beyond dashboards into planning logic.</p>
              </div>
            </div>
            <div className={styles.workflowItem}>
              <span>02</span>
              <div>
                <strong>Business-facing BI</strong>
                <p>Power BI, Tableau, and reporting-oriented projects expand the visual communication layer of the portfolio.</p>
              </div>
            </div>
            <div className={styles.workflowItem}>
              <span>03</span>
              <div>
                <strong>Analytical range</strong>
                <p>Neo4j, MongoDB, and optimization tasks widen the signal without diluting the premium feel of the site.</p>
              </div>
            </div>
          </div>
        </article>
      </section>

      <section className={styles.resourceShelf} id="archive">
        <div className={styles.resourceHeading}>
          <span className={styles.eyebrow}>Full archive</span>
          <h2>Additional analytics projects and downloads.</h2>
          <p>
            The archive below keeps the project set broad while the main page remains curated.
            Each item points to a downloadable artifact or an internal case brief so the full
            analytics footprint stays visible without broken links.
          </p>
        </div>

        <div className={styles.resourceGrid}>
          {analyticsArchive.map((item, index) => (
            <motion.article
              key={item.title}
              className={styles.resourceCard}
              initial={{ opacity: 0, y: 18 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, amount: 0.3 }}
              transition={{ duration: 0.45, delay: index * 0.04 }}
            >
              <span className={styles.resourceTag}>{item.tag}</span>
              <h3>{item.title}</h3>
              <p>{item.summary}</p>

              <div className={styles.resourcePoints}>
                <div className={styles.resourcePoint}>{item.format}</div>
                <div className={styles.resourcePoint}>Public portfolio artifact</div>
                <div className={styles.resourcePoint}>Integrated into the site without changing the main brand direction</div>
              </div>

              <div className={styles.resourceActions}>
                <a
                  href={item.href}
                  target={item.format === "PDF" ? "_blank" : undefined}
                  rel="noreferrer"
                >
                  {item.hrefLabel ?? (item.format === "PDF" ? "Open file" : `Download ${item.format}`)}
                </a>
              </div>
            </motion.article>
          ))}
        </div>
      </section>
    </main>
  );
}
