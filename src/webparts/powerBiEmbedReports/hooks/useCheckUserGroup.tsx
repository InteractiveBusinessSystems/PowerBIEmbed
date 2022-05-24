import {getGraph} from '../config/PNPjsPresets';
import { graphfi, GraphFI } from '@pnp/graph';

export const useCheckUserGroup = (group, currentUserGroups)=> {
  let gName = group.Name;
  let groupName = gName.substring(14);
  const graph: GraphFI = getGraph();
  let contains = false;

  // const currentUserGroups:any = await graphfi(graph).me.getMemberGroups(true);

  currentUserGroups.forEach((userGroup)=>{
      if(groupName.toLowerCase() === userGroup.toLowerCase()){
          contains = true;
      }
  });
  if(contains){
    return true;
  }
  else {
    return false;
  }
}

