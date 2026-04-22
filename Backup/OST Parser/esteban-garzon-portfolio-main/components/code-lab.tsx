"use client";

import { motion } from "framer-motion";
import { useMemo, useState } from "react";
import {
  codeLabGuides,
  codeLabRequirements,
  codeLabSignals,
  codeLabTracks,
  codeLabWorkflow,
  codeSamples,
  profile,
  resources,
} from "@/lib/portfolio-data";
import styles from "./code-lab.module.css";

export function CodeLab() {
  const [activeId, setActiveId] = useState(codeSamples[0]?.id ?? "");

  const activeSample = useMemo(
    () => codeSamples.find((sample) => sample.id === activeId) ?? codeSamples[0],
    [activeId]
  );

  return (
    <main className={styles.page}>
      <div className={styles.backdrop} />

      <header className={styles.header}>
        <a href="/" className={styles.brand}>
          {profile.name}
        </a>

        <div className={styles.headerLinks}>
          <a href="/">Home</a>
          <a href="/analytics">Analytics</a>
          <a href="/#resources">Resources</a>
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
          <span className={styles.eyebrow}>Code Lab</span>
          <h1>Selective code access, architecture context, step-by-step usage, and sanitized engineering proof.</h1>
          <p>
            This page intentionally exposes curated slices of the work rather than a
            full repository dump. The goal is to show engineering thinking,
            integration breadth, execution flow, and practical usage guidance
            without over-sharing credentials, internal identifiers, or raw business data.
          </p>
        </motion.div>
      </section>

      <section className={styles.signalStrip}>
        {codeLabSignals.map((signal, index) => (
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
        {codeLabTracks.map((track, index) => (
          <motion.article
            key={track.title}
            className={styles.trackCard}
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true, amount: 0.3 }}
            transition={{ duration: 0.45, delay: index * 0.08 }}
          >
            <h2>{track.title}</h2>
            <p>{track.summary}</p>
            <div className={styles.trackPoints}>
              {track.points.map((point) => (
                <div key={point} className={styles.trackPoint}>
                  {point}
                </div>
              ))}
            </div>
          </motion.article>
        ))}
      </section>

      <section className={styles.dualSection}>
        <article className={styles.panelCard}>
          <div className={styles.panelHeading}>
            <span className={styles.eyebrow}>Requirements</span>
            <h2>What you need before using the public bundles.</h2>
          </div>

          <div className={styles.requirementGrid}>
            {codeLabRequirements.map((item) => (
              <div key={item.title} className={styles.requirementCard}>
                <strong>{item.title}</strong>
                <p>{item.detail}</p>
              </div>
            ))}
          </div>
        </article>

        <article className={styles.panelCard}>
          <div className={styles.panelHeading}>
            <span className={styles.eyebrow}>Workflow</span>
            <h2>How to use this material step by step.</h2>
          </div>

          <div className={styles.workflowList}>
            {codeLabWorkflow.map((item) => (
              <div key={item.step} className={styles.workflowItem}>
                <span>{item.step}</span>
                <div>
                  <strong>{item.title}</strong>
                  <p>{item.detail}</p>
                </div>
              </div>
            ))}
          </div>
        </article>
      </section>

      <section className={styles.guideSection}>
        <div className={styles.resourceHeading}>
          <span className={styles.eyebrow}>Playbooks</span>
          <h2>Bundle-specific instructions with prerequisites, steps, and expected outputs.</h2>
        </div>

        <div className={styles.guideGrid}>
          {codeLabGuides.map((guide, index) => (
            <motion.article
              key={guide.title}
              className={styles.guideCard}
              initial={{ opacity: 0, y: 18 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, amount: 0.3 }}
              transition={{ duration: 0.45, delay: index * 0.08 }}
            >
              <span className={styles.resourceTag}>{guide.audience}</span>
              <h3>{guide.title}</h3>
              <p>{guide.summary}</p>

              <div className={styles.guideColumns}>
                <div>
                  <strong>Requirements</strong>
                  <div className={styles.guideList}>
                    {guide.requirements.map((item) => (
                      <div key={item} className={styles.guideListItem}>
                        {item}
                      </div>
                    ))}
                  </div>
                </div>

                <div>
                  <strong>Step by step</strong>
                  <div className={styles.guideList}>
                    {guide.steps.map((item) => (
                      <div key={item} className={styles.guideListItem}>
                        {item}
                      </div>
                    ))}
                  </div>
                </div>

                <div>
                  <strong>Expected outputs</strong>
                  <div className={styles.guideList}>
                    {guide.outputs.map((item) => (
                      <div key={item} className={styles.guideListItem}>
                        {item}
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </motion.article>
          ))}
        </div>
      </section>

      <section className={styles.layout}>
        <aside className={styles.sidebar}>
          <div className={styles.sidebarTop}>
            <span className={styles.fileLabel}>Sample Index</span>
            <strong>{codeSamples.length} curated excerpts</strong>
          </div>

          {codeSamples.map((sample, index) => {
            const isActive = sample.id === activeSample.id;
            return (
              <button
                key={sample.id}
                type="button"
                className={`${styles.sampleButton} ${isActive ? styles.sampleButtonActive : ""}`}
                onClick={() => setActiveId(sample.id)}
              >
                <span className={styles.sampleIndex}>
                  {String(index + 1).padStart(2, "0")}
                </span>
                <strong>{sample.title}</strong>
                <small>{sample.category}</small>
              </button>
            );
          })}
        </aside>

        <motion.section
          key={activeSample.id}
          className={styles.viewer}
          initial={{ opacity: 0, y: 18 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.35 }}
        >
          <div className={styles.viewerMeta}>
            <div>
              <span className={styles.fileLabel}>Source</span>
              <h2>{activeSample.title}</h2>
              <p>{activeSample.summary}</p>
            </div>
            <div className={styles.filePill}>{activeSample.file}</div>
          </div>

          <div className={styles.notesGrid}>
            {activeSample.notes.map((note) => (
              <div key={note} className={styles.noteCard}>
                {note}
              </div>
            ))}
          </div>

          <div className={styles.codeFrame}>
            <div className={styles.codeTop}>
              <span />
              <span />
              <span />
              <strong>Read-only excerpt</strong>
            </div>
            <pre className={styles.codeBlock}>
              <code>{activeSample.snippet}</code>
            </pre>
          </div>
        </motion.section>
      </section>

      <section className={styles.resourceShelf}>
        <div className={styles.resourceHeading}>
          <span className={styles.eyebrow}>Public Bundles</span>
          <h2>Downloads that stay useful without exposing internals.</h2>
        </div>

        <div className={styles.resourceGrid}>
          {resources.map((resource, index) => (
            <motion.article
              key={resource.title}
              className={styles.resourceCard}
              initial={{ opacity: 0, y: 18 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, amount: 0.3 }}
              transition={{ duration: 0.45, delay: index * 0.08 }}
            >
              <span className={styles.resourceTag}>{resource.tag}</span>
              <h3>{resource.title}</h3>
              <p>{resource.summary}</p>

              <div className={styles.resourcePoints}>
                {resource.points.map((point) => (
                  <div key={point} className={styles.resourcePoint}>
                    {point}
                  </div>
                ))}
              </div>

              <div className={styles.resourceActions}>
                <a href={resource.downloadHref}>Download package</a>
                {resource.repoHref ? (
                  <a href={resource.repoHref} target="_blank" rel="noreferrer">
                    Open repository
                  </a>
                ) : null}
              </div>
            </motion.article>
          ))}
        </div>
      </section>
    </main>
  );
}
