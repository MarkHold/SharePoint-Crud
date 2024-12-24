import * as React from "react";
import type { IBroadcastandMessagesProps } from "./IBroadcastandMessagesProps";
import { SPFI } from "@pnp/sp";
import { useEffect, useRef, useState } from "react";
import { getSP } from "../../../pnpjsConfig";

// Import your newly created service methods
import { FAQListItem, getFAQItems, addFAQItem, deleteFAQItem } from "../services/sp";

import styles from "./BroadcastandMessages.module.scss";

// (Optional) Graph method if you still need user groups
import { getCurrentUserGroups } from "../services/graph";

// Fluent UI icon
import { FontIcon } from "@fluentui/react/lib/Icon";

const Faq = (props: IBroadcastandMessagesProps) => {

  // SP context
  const _sp: SPFI = getSP(props.context);

  // Accordion items
  const [faqItems, setFaqItems] = useState<FAQListItem[]>([]);
  const [openIndex, setOpenIndex] = useState<number | null>(null);

  // Modal states
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [newTitle, setNewTitle] = useState<string>("");
  const [newDescription, setNewDescription] = useState<string>("");

  /**
   * On mount, get user groups (optional) and load all FAQ items from SharePoint.
   */
  useEffect(() => {
    props.context.msGraphClientFactory.getClient("3").then(async (client) => {
      // Optional: fetch user groups
      const groups = await getCurrentUserGroups(client);
      console.log("User groups from Graph:", groups.value);

      // Fetch items for the accordion
      await refreshFAQItems();
    });
  }, []);

  /**
   * Helper to refresh the FAQ items from SharePoint
   */
  const refreshFAQItems = async () => {
    try {
      const items = await getFAQItems(_sp);
      setFaqItems(items);
    } catch (err) {
      console.error("Error refreshing FAQ items:", err);
    }
  };

  /**
   * Toggle open/close a specific accordion item
   */
  const toggleAccordion = (index: number) => {
    setOpenIndex(openIndex === index ? null : index);
  };

  /**
   * Modal controls
   */
  const openModal = () => setIsModalOpen(true);
  const closeModal = () => {
    setIsModalOpen(false);
    // (Optional) clear form fields on close
    setNewTitle("");
    setNewDescription("");
  };

  /**
   * Submit a new item to the "Prompts" list
   */
  const handleSubmit = async () => {
    try {
      await addFAQItem(_sp, {
        Title: newTitle,
        Description: newDescription,
      });

      // Clear fields & close modal
      setNewTitle("");
      setNewDescription("");
      setIsModalOpen(false);

      alert("New item has been added successfully!");

      // Refresh to show the newly added item
      await refreshFAQItems();
    } catch (error) {
      console.error("Error adding item:", error);
      alert("Error adding item");
    }
  };

  /**
   * Delete an item from "Prompts" by ID, then refresh.
   */
  const handleDeleteFAQItem = async (itemId: number) => {
    try {
      await deleteFAQItem(_sp, itemId);
      alert("Item deleted successfully!");

      // Refresh to remove the deleted item from the accordion
      await refreshFAQItems();
    } catch (error) {
      console.error("Error deleting item:", error);
      alert("Error deleting item");
    }
  };

  return (
    <div className={styles["accordion-container"]}>
      
      {/* Header row with a title on the left and 'New Item' button on the right */}
      <div className={styles["accordion-header-row"]}>
        <h2>FAQ</h2>
        <button className={styles["new-item-button"]} onClick={openModal}>
          New Item
        </button>
      </div>

      {/* If there's an open accordion item, display it at the top */}
      {openIndex !== null && (
        <FaqItem
          faqItem={faqItems[openIndex]}
          key={faqItems[openIndex].ID}
          isOpen={true}
          onClick={() => toggleAccordion(openIndex)}
          isFullWidth={true}
          onDeleteItem={handleDeleteFAQItem}
        />
      )}

      {/* Render the rest of the accordion, skipping the expanded one */}
      {faqItems.map((faqItem, index) => {
        if (index === openIndex) return null; // skip the expanded item
        return (
          <FaqItem
            faqItem={faqItem}
            key={faqItem.ID}
            isOpen={false}
            onClick={() => toggleAccordion(index)}
            isFullWidth={false}
            onDeleteItem={handleDeleteFAQItem}
          />
        );
      })}

      {/* Modal for adding a new item */}
      {isModalOpen && (
        <div className={styles["modal-overlay"]}>
          <div className={styles["modal-content"]}>
            <span className={styles["modal-close"]} onClick={closeModal}>
              &times;
            </span>

            <h3>Add New Prompt</h3>

            <div>
              <label>
                Title:&nbsp;
                <input
                  type="text"
                  value={newTitle}
                  onChange={(e) => setNewTitle(e.target.value)}
                />
              </label>
            </div>

            <div>
              <label>
                Description:&nbsp;
                <textarea
                  value={newDescription}
                  onChange={(e) => setNewDescription(e.target.value)}
                />
              </label>
            </div>

            <button onClick={handleSubmit}>Submit</button>
          </div>
        </div>
      )}
    </div>
  );
};

/**
 * A separate component for each accordion item,
 * including a Delete button in the bottom-right corner.
 */
const FaqItem = (props: {
  faqItem: FAQListItem;
  isOpen: boolean;
  onClick: () => void;
  isFullWidth: boolean;
  onDeleteItem: (id: number) => void;
}) => {
  const { faqItem, isOpen, onClick, isFullWidth, onDeleteItem } = props;
  const contentRef = useRef<HTMLDivElement>(null);

  // Animate open/close
  useEffect(() => {
    if (contentRef.current) {
      if (isOpen) {
        contentRef.current.style.height = `${contentRef.current.scrollHeight}px`;
      } else {
        contentRef.current.style.height = "0px";
      }
    }
  }, [isOpen]);

  return (
    <div
      className={`${styles["accordion-tab"]} ${
        isFullWidth ? styles.fullWidth : ""
      } ${isOpen ? styles.active : ""}`}
    >
      {/* Header area (click to expand/collapse) */}
      <div className={styles["accordion-header"]} onClick={onClick}>
        <div className={styles["accordion-title-container"]}>
          <strong className={styles["accordion-title"]}>{faqItem.Title}</strong>
        </div>
        <span className={styles["accordion-icon"]}>
          <FontIcon
            aria-label="Accordion Toggle"
            iconName={isOpen ? "ChevronUpMed" : "ChevronDownMed"}
          />
        </span>
      </div>

      {/* Content area (collapsible) */}
      <div
        ref={contentRef}
        className={styles["accordion-content"]}
        style={{
          height: "0px",
          overflow: "hidden",
          transition: "height 0.5s ease-in-out",
          position: "relative",
        }}
      >
        {faqItem.Description && <p>{faqItem.Description}</p>}

        {/* Delete button in bottom-right corner */}
        <div style={{ textAlign: "right", marginTop: "10px" }}>
          <button onClick={() => onDeleteItem(faqItem.ID)}>Delete</button>
        </div>
      </div>
    </div>
  );
};

export default Faq;
