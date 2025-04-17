import * as React from 'react';
import styles from './UserModal.module.scss';
import { IUserModalDialogProps } from './IUserModalDialogProps';
import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { escape } from '@microsoft/sp-lodash-subset';

const UserModalDialog: React.FC<IUserModalDialogProps> = (props) => {
  const { isOpen, onDismiss, userData, isDarkTheme } = props;

  if (!userData) {
    return null;
  }

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onDismiss}
      isBlocking={false}
      containerClassName={`${styles.modalContainer} ${isDarkTheme ? styles.dark : ''}`}
    >
      <div className={styles.modalHeader}>
        <IconButton
          className={styles.closeButton}
          iconProps={{ iconName: 'Cancel' }}
          ariaLabel="Close"
          onClick={onDismiss}
        />
      </div>
      <div className={styles.modalContent}>
        <div className={styles.userHeader}>
          <Persona
            imageUrl={userData.photoUrl}
            size={PersonaSize.size100}
            text={userData.title}
            secondaryText={userData.position}
            tertiaryText={userData.email}
          />
        </div>
        
        <div className={styles.modalSection}>
          <h3 className={styles.sectionTitle}>About</h3>
          <p className={styles.sectionText}>{escape(userData.description)}</p>
        </div>
        
        <div className={styles.modalSection}>
          <h3 className={styles.sectionTitle}>Certifications</h3>
          <p className={styles.sectionText}>{escape(userData.certification)}</p>
        </div>
      </div>
    </Modal>
  );
};

export default UserModalDialog;