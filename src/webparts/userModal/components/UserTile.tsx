import * as React from 'react';
import styles from './UserModal.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { IUserTileProps } from './IUserTileProps';
import { Icon } from '@fluentui/react/lib/Icon';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';

const UserTile: React.FC<IUserTileProps> = (props) => {
  const { item, onOpenModal } = props;

  const handleTileClick = (): void => {
    onOpenModal(item);
  };

  return (
    <div 
      className={styles.tileContainer}
      onClick={handleTileClick}
      role="button"
      tabIndex={0}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          handleTileClick();
        }
      }}
    >
      <div className={styles.imageContainer}>
        <Persona
          imageUrl={item.photoUrl}
          size={PersonaSize.size72}
          text={item.title}
          secondaryText={item.position}
        />
      </div>
      <div className={styles.contentContainer}>
        <div>
          <h3 className={styles.title}>{escape(item.title)}</h3>
          <p className={styles.position}>{escape(item.position)}</p>
        </div>
        <div className={styles.arrowIcon}>
          <Icon iconName="ChevronRightMed" />
        </div>
      </div>
    </div>
  );
};

export default UserTile;