import { useCallback, useReducer, useState } from 'react';
import { useGetMemberGroups } from '../hooks/useGetMemberGroups';

export const useCheckMemberGroupsAgainstAGroup = () => {
  const [isAudienced, setIsAudienced] = useState(false);
  const {checkMemberGroups} = useGetMemberGroups();


  const checkMemberGroupsAgainstAGroup = useCallback(async (group) => {
    const gName = group.Name;
    const groupName = gName.substring(14);
    let contains = false;
    const currentUserGroups = await checkMemberGroups();

    currentUserGroups.forEach((userGroup) => {
      if (groupName.toLowerCase() === userGroup.toLowerCase()) {
        contains = true;
      }
    });
    if (contains) {
      setIsAudienced(true);
    }
    else {
      setIsAudienced(false);
    }
    return isAudienced;
  }, []);
  return {checkMemberGroupsAgainstAGroup};

}
