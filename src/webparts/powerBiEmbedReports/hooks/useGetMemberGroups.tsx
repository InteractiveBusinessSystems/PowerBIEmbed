import { useCallback, useState } from 'react';
import { GraphFI, graphfi } from '@pnp/graph';
import { getGraph } from '../config/PNPjsPresets';

export const useGetMemberGroups = () => {
  const [currentUserGroups,setCurrentUserGroups] = useState([]);

  const checkMemberGroups = useCallback(async () => {
    const graph: GraphFI = getGraph();
    const groups:any = await graphfi(graph).me.getMemberGroups(true);


  if(groups){
      setCurrentUserGroups(groups);
  }
  console.log(currentUserGroups);
  return currentUserGroups;
  },[]);
  return {checkMemberGroups};
}
