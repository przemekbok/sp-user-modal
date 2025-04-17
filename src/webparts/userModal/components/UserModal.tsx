import * as React from 'react';
import styles from './UserModal.module.scss';
import { IUserModalProps } from './IUserModalProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IUserItem } from '../UserModalWebPart';
import UserTile from './UserTile';
import UserModalDialog from './UserModalDialog';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';

export default class UserModal extends React.Component<IUserModalProps, {
  currentPage: number;
  isModalOpen: boolean;
  selectedUser: IUserItem | null;
}> {
  
  constructor(props: IUserModalProps) {
    super(props);
    this.state = {
      currentPage: 0,
      isModalOpen: false,
      selectedUser: null
    };
  }

  public render(): React.ReactElement<IUserModalProps> {
    const { 
      userItems, 
      isLoading, 
      itemsPerPage,
      hasTeamsContext,
      isDarkTheme
    } = this.props;

    const { currentPage, isModalOpen, selectedUser } = this.state;
    
    // Calculate total pages
    const totalPages = Math.ceil(userItems.length / itemsPerPage);
    
    // Get items for current page
    const startIndex = currentPage * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const currentItems = userItems.slice(startIndex, endIndex);
    
    // Generate grid class based on items per page
    let gridClass = styles.gridOne;
    if (itemsPerPage === 2) {
      gridClass = styles.gridTwo;
    } else if (itemsPerPage === 3) {
      gridClass = styles.gridThree;
    } else if (itemsPerPage === 4) {
      gridClass = styles.gridFour;
    }

    // Handle navigation
    const goToPreviousPage = (): void => {
      if (currentPage > 0) {
        this.setState({ currentPage: currentPage - 1 });
      }
    };

    const goToNextPage = (): void => {
      if (currentPage < totalPages - 1) {
        this.setState({ currentPage: currentPage + 1 });
      }
    };

    // Handle modal
    const openModal = (user: IUserItem): void => {
      this.setState({ 
        isModalOpen: true,
        selectedUser: user
      });
    };

    const dismissModal = (): void => {
      this.setState({ isModalOpen: false });
    };

    return (
      <div className={`${styles.userModal} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.container}>
          {isLoading ? (
            <div className={styles.spinner}>
              <Spinner size={SpinnerSize.large} label="Loading team members..." />
            </div>
          ) : userItems.length === 0 ? (
            <div className={styles.noItems}>
              <p>No team members found. Please check the list configuration in the web part properties.</p>
            </div>
          ) : (
            <div className={styles.carouselContainer}>
              <div className={`${styles.tilesGrid} ${gridClass}`}>
                {currentItems.map((item: IUserItem) => (
                  <UserTile 
                    key={item.id} 
                    item={item}
                    onOpenModal={openModal}
                    context={this.props.context}
                  />
                ))}
              </div>
              
              {totalPages > 1 && (
                <div className={styles.navigationControls}>
                  <button 
                    className={`${styles.navButton} ${currentPage === 0 ? styles.disabled : ''}`}
                    onClick={goToPreviousPage}
                    disabled={currentPage === 0}
                    aria-label="Previous page"
                  >
                    <Icon iconName="ChevronLeftMed" />
                  </button>
                  <div className={styles.pageIndicator}>
                    {`${currentPage + 1} / ${totalPages}`}
                  </div>
                  <button 
                    className={`${styles.navButton} ${currentPage === totalPages - 1 ? styles.disabled : ''}`}
                    onClick={goToNextPage}
                    disabled={currentPage === totalPages - 1}
                    aria-label="Next page"
                  >
                    <Icon iconName="ChevronRightMed" />
                  </button>
                </div>
              )}
            </div>
          )}

          {isModalOpen && selectedUser && (
            <UserModalDialog 
              isOpen={isModalOpen}
              onDismiss={dismissModal}
              userData={selectedUser}
              isDarkTheme={isDarkTheme}
            />
          )}
        </div>
      </div>
    );
  }
}