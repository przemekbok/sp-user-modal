import * as React from 'react';
import styles from './UserModal.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { IUserTileProps } from './IUserTileProps';
import { Icon } from '@fluentui/react/lib/Icon';

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
        <img src={item.photoUrl} alt={item.title} />
      </div>
      <div className={styles.contentContainer}>
        <h3 className={styles.title}>{escape(item.title)}</h3>
        <p className={styles.position}>{escape(item.position)}</p>
        <div className={styles.arrowIcon}>
          <span>→</span>
        </div>
      </div>
    </div>
  );
};

export default UserTile;