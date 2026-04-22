"use client";

import { motion, useScroll, useSpring, useTransform } from "framer-motion";
import { CSSProperties, useEffect, useState } from "react";
import {
  analyticsFeatured,
  capabilities,
  metrics,
  profile,
  projects,
  resources,
  resumeHighlights,
  resumeItems,
  type Project,
} from "@/lib/portfolio-data";
import styles from "./portfolio-shell.module.css";

const reveal = {
  hidden: { opacity: 0, y: 38 },
  visible: { opacity: 1, y: 0 },
};

function MagneticLink({
  href,
  children,
  variant = "primary",
  target,
  rel,
  download,
}: {
  href: string;
  children: React.ReactNode;
  variant?: "primary" | "ghost";
  target?: string;
  rel?: string;
  download?: boolean | string;
}) {
  const [offset, setOffset] = useState({ x: 0, y: 0 });

  return (
    <motion.a
      href={href}
      target={target}
      rel={rel}
      download={download}
      className={variant === "primary" ? styles.primaryCta : styles.ghostCta}
      onMouseMove={(event) => {
        const bounds = event.currentTarget.getBoundingClientRect();
        const x = event.clientX - (bounds.left + bounds.width / 2);
        const y = event.clientY - (bounds.top + bounds.height / 2);
        setOffset({ x: x * 0.08, y: y * 0.08 });
      }}
      onMouseLeave={() => setOffset({ x: 0, y: 0 })}
      animate={{ x: offset.x, y: offset.y }}
      transition={{ type: "spring", stiffness: 280, damping: 18, mass: 0.35 }}
    >
      <span>{children}</span>
    </motion.a>
  );
}

function ProjectCard({ project, index }: { project: Project; index: number }) {
  const [tilt, setTilt] = useState({ x: 0, y: 0 });

  return (
    <motion.article
      className={styles.projectCard}
      initial="hidden"
      whileInView="visible"
      viewport={{ once: true, amount: 0.35 }}
      variants={reveal}
      transition={{ duration: 0.6, delay: index * 0.12 }}
      onMouseMove={(event) => {
        const bounds = event.currentTarget.getBoundingClientRect();
        const px = (event.clientX - bounds.left) / bounds.width;
        const py = (event.clientY - bounds.top) / bounds.height;
        setTilt({
          x: (py - 0.5) * -9,
          y: (px - 0.5) * 12,
        });
      }}
      onMouseLeave={() => setTilt({ x: 0, y: 0 })}
      style={
        {
          transform: `perspective(1200px) rotateX(${tilt.x}deg) rotateY(${tilt.y}deg)`,
          "--project-accent": project.accent,
        } as CSSProperties
      }
    >
      <div className={styles.projectHeader}>
        <span className={styles.projectTag}>{project.tag}</span>
        <span className={styles.projectIndex}>0{index + 1}</span>
      </div>
      <h3>{project.title}</h3>
      <p>{project.summary}</p>
      <div className={styles.projectDivider} />
      <p className={styles.projectOutcome}>{project.outcome}</p>
      <div className={styles.stackRow}>
        {project.stack.map((item) => (
          <span key={item}>{item}</span>
        ))}
      </div>
    </motion.article>
  );
}

export function PortfolioShell() {
  const [cursor, setCursor] = useState({ x: 50, y: 20 });
  const { scrollYProgress } = useScroll();
  const progress = useSpring(scrollYProgress, {
    stiffness: 120,
    damping: 28,
    mass: 0.25,
  });
  const heroLift = useTransform(progress, [0, 0.25], [0, -40]);
  const heroFade = useTransform(progress, [0, 0.25], [1, 0.2]);

  useEffect(() => {
    const handleMove = (event: MouseEvent) => {
      const x = (event.clientX / window.innerWidth) * 100;
      const y = (event.clientY / window.innerHeight) * 100;
      setCursor({ x, y });
    };

    window.addEventListener("mousemove", handleMove);
    return () => window.removeEventListener("mousemove", handleMove);
  }, []);

  return (
    <main
      className={styles.page}
      style={
        {
          "--cursor-x": `${cursor.x}%`,
          "--cursor-y": `${cursor.y}%`,
        } as CSSProperties
      }
    >
      <motion.div className={styles.progressBar} style={{ scaleX: progress }} />
      <div className={styles.spotlight} />
      <div className={styles.gridGlow} />

      <header className={styles.topbar}>
        <a href="#top" className={styles.brand}>
          {profile.name}
        </a>

        <nav className={styles.nav}>
          <a href="#work">Work</a>
          <a href="#analytics">Analytics</a>
          <a href="#resources">Resources</a>
          <a href="#resume">Resume</a>
          <a href="/code">Code Lab</a>
          <a href="/analytics">Analytics Lab</a>
          <a href="#contact">Contact</a>
          <a href={profile.linkedin} target="_blank" rel="noreferrer">
            LinkedIn
          </a>
        </nav>
      </header>

      <section className={styles.hero} id="top">
        <motion.div
          className={styles.heroCopy}
          style={{ y: heroLift, opacity: heroFade }}
          initial={{ opacity: 0, y: 30 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.8 }}
        >
          <span className={styles.eyebrow}>{profile.role}</span>
          <h1>{profile.heroTitle}</h1>
          <p>{profile.heroText}</p>

          <div className={styles.ctaRow}>
            <MagneticLink href="#work">Explore work</MagneticLink>
            <MagneticLink href={profile.resumeHref} variant="ghost" target="_blank" rel="noreferrer" download>
              Download resume
            </MagneticLink>
            <MagneticLink href="/code" variant="ghost">
              Open Code Lab
            </MagneticLink>
            <MagneticLink href={profile.linkedin} variant="ghost" target="_blank" rel="noreferrer">
              LinkedIn
            </MagneticLink>
          </div>
        </motion.div>

        <motion.div
          className={styles.heroPanel}
          initial={{ opacity: 0, scale: 0.95, y: 24 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          transition={{ duration: 0.9, delay: 0.15 }}
        >
          <div className={styles.panelTop}>
            <span className={styles.liveDot} />
            <span>System Signal</span>
          </div>
          <div className={styles.orbitField}>
            <div className={styles.orbitOne} />
            <div className={styles.orbitTwo} />
            <div className={styles.orbitThree} />
            <div className={styles.coreBadge}>Modern Ops</div>
          </div>
          <div className={styles.metricList}>
            {metrics.map((metric) => (
              <div key={metric.label} className={styles.metricItem}>
                <span>{metric.label}</span>
                <strong>{metric.value}</strong>
              </div>
            ))}
          </div>
        </motion.div>
      </section>

      <section className={styles.marqueeSection} aria-label="technologies">
        <div className={styles.marqueeTrack}>
          <span>Python</span>
          <span>Next.js</span>
          <span>PostgreSQL</span>
          <span>SQL Server</span>
          <span>Automation</span>
          <span>ETL</span>
          <span>Data Reliability</span>
          <span>API Integrations</span>
          <span>Systems Design</span>
          <span>Interactive UI</span>
        </div>
      </section>

      <section className={styles.section} id="work">
        <motion.div
          className={styles.sectionHeading}
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, amount: 0.35 }}
          variants={reveal}
          transition={{ duration: 0.6 }}
        >
          <span>Selected Work</span>
          <h2>Real-world systems, presented with a product-grade visual layer.</h2>
        </motion.div>

        <div className={styles.projectsGrid}>
          {projects.map((project, index) => (
            <ProjectCard key={project.title} project={project} index={index} />
          ))}
        </div>
      </section>

      <section className={styles.section} id="analytics">
        <motion.div
          className={styles.sectionHeading}
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, amount: 0.35 }}
          variants={reveal}
          transition={{ duration: 0.6 }}
        >
          <span>Applied Analytics</span>
          <h2>A separate analytics layer that adds forecasting, BI, visualization, and graph-driven case work.</h2>
        </motion.div>

        <div className={styles.projectsGrid}>
          {analyticsFeatured.slice(0, 3).map((project, index) => (
            <ProjectCard key={project.title} project={project} index={index} />
          ))}
        </div>

        <div className={styles.ctaRow}>
          <MagneticLink href="/analytics">Open Analytics Portfolio</MagneticLink>
          <MagneticLink href={analyticsFeatured[0].href} variant="ghost">
            {analyticsFeatured[0].hrefLabel}
          </MagneticLink>
        </div>
      </section>

      <section className={`${styles.section} ${styles.processSection}`}>
        <motion.div
          className={styles.sectionHeading}
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, amount: 0.35 }}
          variants={reveal}
          transition={{ duration: 0.6 }}
        >
          <span>Build Philosophy</span>
          <h2>I care about systems that work under pressure and still feel refined.</h2>
        </motion.div>

        <div className={styles.processGrid}>
          <motion.div
            className={styles.processCard}
            initial="hidden"
            whileInView="visible"
            viewport={{ once: true, amount: 0.3 }}
            variants={reveal}
            transition={{ duration: 0.55, delay: 0.05 }}
          >
            <strong>01</strong>
            <h3>Signal first</h3>
            <p>
              I start by finding where data breaks down and which decision the
              software must make easier for the team.
            </p>
          </motion.div>

          <motion.div
            className={styles.processCard}
            initial="hidden"
            whileInView="visible"
            viewport={{ once: true, amount: 0.3 }}
            variants={reveal}
            transition={{ duration: 0.55, delay: 0.15 }}
          >
            <strong>02</strong>
            <h3>Reliability by design</h3>
            <p>
              Deduplication, traceability, fallbacks, and safe writes keep the
              automation resilient in real environments.
            </p>
          </motion.div>

          <motion.div
            className={styles.processCard}
            initial="hidden"
            whileInView="visible"
            viewport={{ once: true, amount: 0.3 }}
            variants={reveal}
            transition={{ duration: 0.55, delay: 0.25 }}
          >
            <strong>03</strong>
            <h3>Presentation matters</h3>
            <p>
              Internal engineering work still deserves interfaces that are clear,
              elegant, and satisfying to explore.
            </p>
          </motion.div>
        </div>
      </section>

      <section className={`${styles.section} ${styles.capabilitiesSection}`}>
        <motion.div
          className={styles.sectionHeading}
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, amount: 0.35 }}
          variants={reveal}
          transition={{ duration: 0.6 }}
        >
          <span>Capabilities</span>
          <h2>Strong backend foundations, expressive visuals, and operational focus.</h2>
        </motion.div>

        <div className={styles.capabilitiesGrid}>
          {capabilities.map((item, index) => (
            <motion.div
              key={item}
              className={styles.capabilityPill}
              initial={{ opacity: 0, y: 24 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, amount: 0.3 }}
              transition={{ duration: 0.45, delay: index * 0.05 }}
            >
              {item}
            </motion.div>
          ))}
        </div>
      </section>

      <section className={`${styles.section} ${styles.resourcesSection}`} id="resources">
        <motion.div
          className={styles.sectionHeading}
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, amount: 0.35 }}
          variants={reveal}
          transition={{ duration: 0.6 }}
        >
          <span>Resources & Contributions</span>
          <h2>Useful public downloads, packaged cleanly and safe to share.</h2>
        </motion.div>

        <div className={styles.resourcesGrid}>
          {resources.map((resource, index) => (
            <motion.article
              key={resource.title}
              className={styles.resourceCard}
              initial={{ opacity: 0, y: 24 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, amount: 0.3 }}
              transition={{ duration: 0.5, delay: index * 0.08 }}
            >
              <div className={styles.resourceTop}>
                <span className={styles.resourceTag}>{resource.tag}</span>
                <h3>{resource.title}</h3>
                <p>{resource.summary}</p>
              </div>

              <div className={styles.resourcePoints}>
                {resource.points.map((point) => (
                  <div key={point} className={styles.resourcePoint}>
                    {point}
                  </div>
                ))}
              </div>

              <div className={styles.resourceActions}>
                <MagneticLink href={resource.downloadHref}>Download source package</MagneticLink>
                {resource.repoHref ? (
                  <MagneticLink href={resource.repoHref} variant="ghost" target="_blank" rel="noreferrer">
                    View repository
                  </MagneticLink>
                ) : null}
              </div>
            </motion.article>
          ))}
        </div>
      </section>

      <section className={`${styles.section} ${styles.resumeSection}`} id="resume">
        <motion.div
          className={styles.sectionHeading}
          initial="hidden"
          whileInView="visible"
          viewport={{ once: true, amount: 0.35 }}
          variants={reveal}
          transition={{ duration: 0.6 }}
        >
          <span>Resume Snapshot</span>
          <h2>A concise professional summary, not just a project gallery.</h2>
        </motion.div>

        <div className={styles.resumeGrid}>
          <motion.article
            className={styles.resumeSummaryCard}
            initial={{ opacity: 0, y: 24 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true, amount: 0.3 }}
            transition={{ duration: 0.55 }}
          >
            <span className={styles.resumeLabel}>Professional Summary</span>
            <h3>{profile.role}</h3>
            <p>{profile.resumeSummary}</p>

            <div className={styles.ctaRow}>
              <MagneticLink href={profile.resumeHref} target="_blank" rel="noreferrer" download>
                Download full CV
              </MagneticLink>
            </div>

            <div className={styles.resumeHighlightList}>
              {resumeHighlights.map((item) => (
                <div key={item} className={styles.resumeHighlight}>
                  {item}
                </div>
              ))}
            </div>
          </motion.article>

          <div className={styles.resumeTimeline}>
            {resumeItems.map((item, index) => (
              <motion.article
                key={item.title}
                className={styles.resumeItem}
                initial={{ opacity: 0, y: 24 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true, amount: 0.3 }}
                transition={{ duration: 0.5, delay: index * 0.08 }}
              >
                <span className={styles.resumePeriod}>{item.period}</span>
                <h3>{item.title}</h3>
                <p>{item.description}</p>
              </motion.article>
            ))}
          </div>
        </div>
      </section>

      <section className={styles.contact} id="contact">
        <motion.div
          className={styles.contactPanel}
          initial={{ opacity: 0, y: 30 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true, amount: 0.35 }}
          transition={{ duration: 0.65 }}
        >
          <span className={styles.contactEyebrow}>Contact</span>
          <h2>Let&apos;s build better operational software.</h2>
          <p>
            If you want to talk about automation, ETL systems, internal tools, or
            product-minded engineering, you can reach me directly here.
          </p>

          <div className={styles.contactGrid}>
            <a className={styles.contactCard} href={`mailto:${profile.email}`}>
              <span>Email</span>
              <strong>{profile.email}</strong>
            </a>
            <a className={styles.contactCard} href={`tel:${profile.phone.replace(/\s+/g, "")}`}>
              <span>Phone</span>
              <strong>{profile.phone}</strong>
            </a>
            <a
              className={styles.contactCard}
              href={profile.linkedin}
              target="_blank"
              rel="noreferrer"
            >
              <span>LinkedIn</span>
              <strong>View profile</strong>
            </a>
          </div>

          <div className={styles.contactActions}>
            <MagneticLink href={`mailto:${profile.email}`}>Email me</MagneticLink>
            <MagneticLink href={profile.resumeHref} variant="ghost" target="_blank" rel="noreferrer" download>
              Download resume
            </MagneticLink>
            <MagneticLink href="/code" variant="ghost">
              View limited code access
            </MagneticLink>
            <MagneticLink href="#top" variant="ghost">
              Back to top
            </MagneticLink>
          </div>
        </motion.div>
      </section>
    </main>
  );
}
